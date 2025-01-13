function TFIDFSearch(documents, options) {
  options = options || {};
  this.documents = documents;
  this.docCount = documents.length;
  this.useStopWords =
    options.useStopWords !== undefined ? options.useStopWords : true;
  this.matchThreshold = options.matchThreshold || 0.0;
  this.minTermSize = options.minTermSize || 3;
  this.wordDocs = {};
  this.timing = {
    processDocuments: 0,
    lastSearch: 0,
  };
  this.processDocuments();
}

// Define stop words as a property of the constructor
TFIDFSearch.STOP_WORDS = {
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
};

TFIDFSearch.prototype.processDocuments = function () {
  var startTime = new Date().getTime();

  for (var i = 0; i < this.documents.length; i++) {
    var words = this.tokenize(this.documents[i].text);
    for (var j = 0; j < words.length; j++) {
      var word = words[j];
      if (!this.wordDocs[word]) {
        this.wordDocs[word] = [];
      }
      if (this.wordDocs[word].indexOf(i) === -1) {
        this.wordDocs[word].push(i);
      }
    }
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

TFIDFSearch.prototype.tf = function (word, docText) {
  var words = this.tokenize(docText);
  var wordCount = words.filter(function (w) {
    return w === word;
  }).length;
  return wordCount / words.length;
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
  var scoreDetails = new Array(this.docCount).fill().map(function () {
    return {
      wordScores: {},
      totalScore: 0,
    };
  });

  for (var i = 0; i < queryWords.length; i++) {
    var word = queryWords[i];
    var idfScore = this.idf(word);

    for (var j = 0; j < this.documents.length; j++) {
      var tfScore = this.tf(word, this.documents[j].text);
      var wordScore = tfScore * idfScore;
      scores[j] += wordScore;

      // Store detailed scoring information
      scoreDetails[j].wordScores[word] = {
        tf: tfScore,
        idf: this.idf(word),
        combined: wordScore,
      };
    }
  }

  // Add term count boost before creating results array
  for (var k = 0; k < scores.length; k++) {
    var matchingTerms = Object.values(scoreDetails[k].wordScores).filter(
      (score) => score.combined > 0
    ).length;
    var termCountBoost = Math.pow(1.2, matchingTerms - 1);
    scores[k] *= termCountBoost;
    scoreDetails[k].totalScore = scores[k]; // Update the total score in details
  }

  var results = [];
  for (var k = 0; k < scores.length; k++) {
    if (scores[k] > this.matchThreshold) {
      results.push({
        id: this.documents[k].id,
        text: this.documents[k].text,
        updatedText: this.documents[k].updatedText,
        score: scores[k],
        details: scoreDetails[k],
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
  startRow = 2
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

  return new TFIDFSearch(documents, {
    useStopWords: true,
    matchThreshold: 0.25,
    minTermSize: 3,
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
  results.forEach(function (result, index) {
    Logger.log("\n--- Result " + (index + 1) + " ---");
    Logger.log("ID: " + result.id);
    Logger.log("Text: " + result.text);
    Logger.log("Date: " + result.date);
    Logger.log("Category: " + result.category);
    Logger.log("Total Score: " + result.score.toFixed(4));

    // Expand the word scores
    Logger.log("Scoring Breakdown:");
    Object.keys(result.details.wordScores).forEach(function (word) {
      var scores = result.details.wordScores[word];
      Logger.log("  Term: '" + word + "'");
      Logger.log("    TF (term frequency): " + scores.tf.toFixed(4));
      Logger.log("    IDF (inverse doc frequency): " + scores.idf.toFixed(4));
      Logger.log("    Combined Score: " + scores.combined.toFixed(4));

      // Calculate percentage contribution to total score
      var contribution = ((scores.combined / result.score) * 100).toFixed(2);
      Logger.log("    Contribution to total score: " + contribution + "%");
    });
  });

  // Add some corpus statistics
  Logger.log("\n=== Corpus Statistics ===");
  Logger.log("Total documents: " + searcher.docCount);

  // Calculate term statistics for query terms
  var queryTerms = searcher.tokenize(query);
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
