function TFIDFSearch(documents, options) {
  options = options || {};
  this.documents = documents;
  this.docCount = documents.length;
  this.useStopWords =
    options.useStopWords !== undefined ? options.useStopWords : true;
  this.matchThreshold = options.matchThreshold || 0.0;
  this.minTermSize = options.minTermSize || 3;
  // Null-prototype, for the same reason as STOP_WORDS: this map is keyed by
  // arbitrary tokens, and "constructor" would otherwise resolve to an inherited
  // property rather than a missing one.
  this.wordDocs = Object.create(null);
  this.timing = {
    processDocuments: 0,
    lastSearch: 0,
  };
  this.processDocuments();
}

// Define stop words as a property of the constructor.
//
// Null-prototype, because this object is looked up by arbitrary tokens from
// transaction text. On a plain object literal the words "constructor" and
// "__proto__" resolve to inherited Object.prototype members, come back truthy,
// and are silently discarded as stop words.
TFIDFSearch.STOP_WORDS = Object.assign(Object.create(null), {
  a: true,
  an: true,
  and: true,
  are: true,
  as: true,
  at: true,
  be: true,
  by: true,
  for: true,
  from: true,
  has: true,
  he: true,
  in: true,
  is: true,
  it: true,
  its: true,
  of: true,
  on: true,
  that: true,
  the: true,
  to: true,
  was: true,
  were: true,
  will: true,
  with: true,
});

TFIDFSearch.prototype.processDocuments = function () {
  var startTime = new Date().getTime();

  // Term frequencies are computed once per document here, and search() reads
  // them back. Previously tf() re-tokenised a document's full text on every
  // comparison, so tokenize() ran (query terms x documents) times per search.
  this.docTerms = new Array(this.documents.length);

  for (var i = 0; i < this.documents.length; i++) {
    var words = this.tokenize(this.documents[i].text);
    var counts = Object.create(null);

    for (var j = 0; j < words.length; j++) {
      var word = words[j];

      if (counts[word] === undefined) {
        counts[word] = 0;
        // First time this term appears in this document, so record the document
        // against it exactly once. This replaces the old indexOf membership
        // scan, which was linear in the number of documents already matched.
        if (!this.wordDocs[word]) {
          this.wordDocs[word] = [];
        }
        this.wordDocs[word].push(i);
      }

      counts[word]++;
    }

    this.docTerms[i] = { counts: counts, total: words.length };
  }

  this.timing.processDocuments = new Date().getTime() - startTime;
};

TFIDFSearch.prototype.tokenize = function (text) {
  var cleanText = String(text)
    .toLowerCase()
    .replace(/[.,!?;:'"()\[\]{}""''`#*]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[–—-]+/g, " ")
    .replace(/['']s\b/g, "")
    .replace(/n['']t\b/g, "not")
    .replace(/['']ve\b/g, "have")
    .replace(/['']re\b/g, "are")
    .replace(/['']ll\b/g, "will")
    .replace(/['']d\b/g, "would")
    .trim();

  var self = this;
  var tokens = cleanText.split(" ").filter(function (word) {
    var cleaned = word.replace(/[^\w-]/g, "");
    return cleaned.length >= self.minTermSize;
  });

  return this.useStopWords
    ? tokens.filter(function (word) {
        return !TFIDFSearch.STOP_WORDS[word];
      })
    : tokens;
};

// Term frequency of a term within a document, by document index. Documents that
// tokenise to nothing divide by zero and produce NaN, which then fails the
// matchThreshold comparison and drops out of the results - as it did before.
TFIDFSearch.prototype.tf = function (word, docIndex) {
  var doc = this.docTerms[docIndex];
  return (doc.counts[word] || 0) / doc.total;
};

TFIDFSearch.prototype.idf = function (word) {
  var docsWithWord = (this.wordDocs[word] || []).length;
  if (docsWithWord === 0) return 0;
  return Math.log(this.docCount / docsWithWord);
};

TFIDFSearch.prototype.search = function (query, limit) {
  var startTime = new Date().getTime();
  limit = limit || 5;

  var queryWords = this.tokenize(query);
  var scores = new Array(this.docCount).fill(0);

  // How many DISTINCT query terms matched each document. A term repeated in the
  // query contributes to the score on each occurrence but lifts this count only
  // once; the previous implementation got that behaviour implicitly by keying a
  // per-document map on the term.
  var matchedTermCounts = new Array(this.docCount).fill(0);
  var countedWords = Object.create(null);

  for (var i = 0; i < queryWords.length; i++) {
    var word = queryWords[i];
    var idfScore = this.idf(word);
    var isFirstOccurrence = countedWords[word] === undefined;
    countedWords[word] = true;

    for (var j = 0; j < this.docCount; j++) {
      var wordScore = this.tf(word, j) * idfScore;
      scores[j] += wordScore;

      if (isFirstOccurrence && wordScore > 0) {
        matchedTermCounts[j]++;
      }
    }
  }

  var results = [];
  for (var k = 0; k < this.docCount; k++) {
    // Favour documents that matched more of the query's distinct terms.
    var score = scores[k] * Math.pow(1.2, matchedTermCounts[k] - 1);

    if (score > this.matchThreshold) {
      results.push({
        id: this.documents[k].id,
        text: this.documents[k].text,
        updatedText: this.documents[k].updatedText,
        score: score,
        // Index into this.documents, so per-term scoring can be reconstructed
        // on demand by printResultsDebugging without paying for it every search.
        docIndex: k,
        date: this.documents[k].date,
        category: this.documents[k].category,
        amount: this.documents[k].amount,
      });
    }
  }

  // Sort by score first, then by date if scores are equal
  results.sort(function (a, b) {
    if (Math.abs(b.score - a.score) < 0.000001) {
      // Use small epsilon for floating point comparison
      // If dates are available, sort by date
      if (a.date && b.date) {
        return new Date(b.date) - new Date(a.date);
      }
      return 0;
    }
    return b.score - a.score;
  });

  this.timing.lastSearch = new Date().getTime() - startTime;
  return results.slice(0, limit);
};

/**
 * Creates a search index from a sheet range
 * @param {string} sheetName Name of the sheet containing the data
 * @param {string} idColumn Column letter for IDs (e.g., "A")
 * @param {string} textColumn Column letter for text content (e.g., "B")
 * @param {string} updatedTextClumn Column letter for updated text content (e.g., "B")
 * @param {string} dateColumn Column letter for dates (e.g., "C")
 * @param {string} categoryColumn Column letter for category (e.g., "D")
 * @param {string} amountColumn Column letter for amount (e.g., "C")
 * @param {number} startRow First row of data (e.g., 2 to skip header)
 * @returns {TFIDFSearch} Search instance
 */
function createSearchIndex(
  sheetName,
  idColumn,
  textColumn,
  updatedTextColumn,
  dateColumn,
  categoryColumn,
  amountColumn,
  startRow = 2,
  options = {}
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  const idRange = sheet.getRange(
    `${idColumn}${startRow}:${idColumn}${lastRow}`
  );
  const textRange = sheet.getRange(
    `${textColumn}${startRow}:${textColumn}${lastRow}`
  );
  const updatedTextRange = sheet.getRange(
    `${updatedTextColumn}${startRow}:${updatedTextColumn}${lastRow}`
  );
  const dateRange = sheet.getRange(
    `${dateColumn}${startRow}:${dateColumn}${lastRow}`
  );
  const categoryRange = sheet.getRange(
    `${categoryColumn}${startRow}:${categoryColumn}${lastRow}`
  );
  const amountRange = sheet.getRange(
    `${amountColumn}${startRow}:${amountColumn}${lastRow}`
  );

  const ids = idRange.getValues().flat();
  const texts = textRange.getValues().flat();
  const updatedTexts = updatedTextRange.getValues().flat();
  const dates = dateRange.getValues().flat();
  const categories = categoryRange.getValues().flat();
  const amounts = amountRange.getValues().flat();

  const documents = [];

  // Process each row, skipping those without required category
  for (let i = 0; i < ids.length; i++) {
    // Skip if category is empty or undefined
    if (!categories[i]) {
      continue;
    }

    documents.push({
      id: ids[i],
      text: texts[i] || "",
      updatedText: updatedTexts[i] || "",
      date: dates[i] ? new Date(dates[i]) : null,
      category: categories[i],
      amount: amounts[i],
    });
  }

  // Callers pass tuning options through createSearchIndexWithStandardColumns;
  // before this they were silently dropped, because the parameter did not exist.
  return new TFIDFSearch(documents, {
    useStopWords:
      options.useStopWords !== undefined ? options.useStopWords : true,
    matchThreshold:
      options.matchThreshold !== undefined ? options.matchThreshold : 0.25,
    minTermSize: options.minTermSize !== undefined ? options.minTermSize : 3,
  });
}

/**
 * Searches the index and writes results to a sheet
 * @param {TFIDFSearch} searcher Search instance
 * @param {string} query Search query
 * @param {string} outputSheetName Name of sheet to write results
 * @param {number} limit Maximum number of results
 */
function searchAndWriteResults(searcher, query, outputSheetName, limit = 5) {
  const results = searcher.search(query, limit);

  const outputSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(outputSheetName) ||
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(outputSheetName);

  // Clear existing content
  outputSheet.clear();

  // Write headers
  outputSheet
    .getRange("A1:F1")
    .setValues([["ID", "Text", "Score", "Category", "Date", "Amount"]]);

  // Write results
  if (results.length > 0) {
    const resultData = results.map((result) => [
      result.id,
      result.text,
      result.score,
      result.category,
      result.date,
      result.amount,
    ]);
    outputSheet.getRange(2, 1, resultData.length, 6).setValues(resultData);
  }
}

function printResultsDebugging(query, searcher, results) {
  var queryTerms = searcher.tokenize(query);

  results.forEach(function (result, index) {
    Logger.log("\n--- Result " + (index + 1) + " ---");
    Logger.log("ID: " + result.id);
    Logger.log("Text: " + result.text);
    Logger.log("Date: " + result.date);
    Logger.log("Category: " + result.category);
    Logger.log("Total Score: " + result.score.toFixed(4));

    // Rebuilt from the index rather than carried on every result, so ordinary
    // searches do not pay to build a breakdown nothing reads.
    Logger.log("Scoring Breakdown:");
    queryTerms.forEach(function (word) {
      var tf = searcher.tf(word, result.docIndex);
      var idf = searcher.idf(word);
      var combined = tf * idf;

      if (!(combined > 0)) {
        return;
      }

      Logger.log("  Term: '" + word + "'");
      Logger.log("    TF (term frequency): " + tf.toFixed(4));
      Logger.log("    IDF (inverse doc frequency): " + idf.toFixed(4));
      Logger.log("    Combined Score: " + combined.toFixed(4));

      var contribution = ((combined / result.score) * 100).toFixed(2);
      Logger.log("    Contribution to total score: " + contribution + "%");
    });
  });

  // Add some corpus statistics
  Logger.log("\n=== Corpus Statistics ===");
  Logger.log("Total documents: " + searcher.docCount);

  queryTerms.forEach(function (term) {
    var docsWithTerm = (searcher.wordDocs[term] || []).length;
    var frequency = ((docsWithTerm / searcher.docCount) * 100).toFixed(2);
    Logger.log(
      "Term '" +
        term +
        "' appears in " +
        docsWithTerm +
        " documents (" +
        frequency +
        "% of corpus)"
    );
  });
}

function testSearch() {
  var query = "TST* STONEMILL MATCHA SAN FRANCISCOCA";
  var limit = 5;
  var printDebuggingInfo = false;

  // Create search index from sheet data
  var searcher = createSearchIndexWithStandardColumns({
    minTermSize: 3,
  });

  const results = searcher.search(query, limit);
  Logger.log("\n=== Search Results for: '" + query + "' ===");
  Logger.log("Index build time: " + searcher.timing.processDocuments + "ms");
  Logger.log("Search time: " + searcher.timing.lastSearch + "ms");

  if (printDebuggingInfo) {
    printResultsDebugging(query, searcher, results);
  } else {
    console.log(results);
  }
}

/**
 * Example usage in Google Apps Script
 */
function searchFromActiveCell() {
  var activeCell = SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
    .getActiveCell();
  var query = activeCell.getValue();

  // Log for debugging
  Logger.log("Active cell: " + activeCell.getA1Notation());
  Logger.log("Query value: " + query);

  // Create search index from sheet data
  var searcher = createSearchIndexWithStandardColumns({
    minTermSize: 3,
  });

  // Search and write results to new sheet
  searchAndWriteResults(searcher, query, "Search Results", 5);

  // Optional: Show confirmation toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Search completed for: "' + query + '"',
    "Search Status"
  );
}
