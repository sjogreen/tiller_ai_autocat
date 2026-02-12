// API Keys
const OPENAI_API_KEY = '';
const GOOGLE_API_KEY = '';

// LLM To Use
const AI_PROVIDER = 'gemini' // Can be 'gemini' or 'openai'
const GPT_MODEL = 'gpt-4o-mini' // Can be any openai model designator

// Sheet Names
const TRANSACTION_SHEET_NAME = "Transactions";
const CATEGORY_SHEET_NAME = "Categories";

// Column Names
const TRANSACTION_ID_COL_NAME = "Transaction ID";
const ORIGINAL_DESCRIPTION_COL_NAME = "Full Description";
const DESCRIPTION_COL_NAME = "Description";
const CATEGORY_COL_NAME = "Category";
const AI_AUTOCAT_COL_NAME = "AI AutoCat";
const DATE_COL_NAME = "Date";
const AMOUNT_COL_NAME = "Amount";

// Fallback Transaction Category (to be used when we don't know how to categorize a transaction)
const FALLBACK_CATEGORY = "To Be Categorized";

// Other Misc Paramaters
const MAX_BATCH_SIZE = 50;
var TRANSACTION_SEARCHER = null;

function categorizeUncategorizedTransactions() {
  var uncategorizedTransactions = getTransactionsToCategorize();

  var numTxnsToCategorize = uncategorizedTransactions.length;
  if (numTxnsToCategorize == 0) {
    Logger.log("No uncategorized transactions found");
    return;
  }

  Logger.log("Found " + numTxnsToCategorize + " transactions to categorize");
  Logger.log("Looking for historical similar transactions...");

  var transactionList = [];
  for (var i = 0; i < uncategorizedTransactions.length; i++) {
    var similarTransactions = findSimilarTransactions(
      uncategorizedTransactions[i][1]
    );

    transactionList.push({
      transaction_id: uncategorizedTransactions[i][0],
      original_description: uncategorizedTransactions[i][1],
      previous_transactions: similarTransactions,
    });
  }

  Logger.log(
    "Processing this set of transactions and similar transactions:"
  );
  Logger.log(transactionList);

  var categoryList = getAllowedCategories();

  var updatedTransactions;
  if (AI_PROVIDER == 'gemini') {
    Logger.log(
      "Using Gemini"
    );

    updatedTransactions = lookupDescAndCategoryGemini(
      transactionList,
      categoryList
    );
  } else {
    Logger.log(
      "Using OpenAI"
    );

    updatedTransactions = lookupDescAndCategoryOpenai(
      transactionList,
      categoryList
    );
  }
  
  if (updatedTransactions != null) {
    Logger.log(
      "The selected AI returned the following sugested categories and descriptions:"
    );
    Logger.log(updatedTransactions);
    Logger.log("Writing updated transactions into your sheet...");
    writeUpdatedTransactions(updatedTransactions, categoryList);
    Logger.log("Finished updating your sheet!");
  }
}

// Gets all transactions that have an original description but no category set
function getTransactionsToCategorize() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    TRANSACTION_SHEET_NAME
  );
  var headers = sheet.getRange("1:1").getValues()[0];

  var txnIDColLetter = getColumnLetterFromColumnHeader(
    headers,
    TRANSACTION_ID_COL_NAME
  );
  var origDescColLetter = getColumnLetterFromColumnHeader(
    headers,
    ORIGINAL_DESCRIPTION_COL_NAME
  );
  var categoryColLetter = getColumnLetterFromColumnHeader(
    headers,
    CATEGORY_COL_NAME
  );
  var lastColLetter = getColumnLetterFromColumnHeader(
    headers,
    headers[headers.length - 1]
  );

  var queryString =
    "SELECT " +
    txnIDColLetter +
    ", " +
    origDescColLetter +
    " WHERE " +
    origDescColLetter +
    " is not null AND " +
    categoryColLetter +
    " is null LIMIT " +
    MAX_BATCH_SIZE;

  var uncategorizedTransactions = Utils.gvizQuery(
    SpreadsheetApp.getActiveSpreadsheet().getId(),
    queryString,
    TRANSACTION_SHEET_NAME,
    "A:" + lastColLetter
  );

  return uncategorizedTransactions;
}

function createSearchIndexWithStandardColumns(options) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    TRANSACTION_SHEET_NAME
  );
  var headers = sheet.getRange("1:1").getValues()[0];

  var idColLetter = getColumnLetterFromColumnHeader(
    headers,
    TRANSACTION_ID_COL_NAME
  );
  var descColLetter = getColumnLetterFromColumnHeader(
    headers,
    DESCRIPTION_COL_NAME
  );
  var origDescColLetter = getColumnLetterFromColumnHeader(
    headers,
    ORIGINAL_DESCRIPTION_COL_NAME
  );
  var categoryColLetter = getColumnLetterFromColumnHeader(
    headers,
    CATEGORY_COL_NAME
  );
  var dateColLetter = getColumnLetterFromColumnHeader(headers, DATE_COL_NAME);
  var amountColLetter = getColumnLetterFromColumnHeader(
    headers,
    AMOUNT_COL_NAME
  );

  var searcher = createSearchIndex(
    TRANSACTION_SHEET_NAME,
    idColLetter, // ID Column
    origDescColLetter, // text column
    descColLetter, // updated text column
    dateColLetter, // date column (for breaking ranking ties)
    categoryColLetter, // category column
    amountColLetter, // amount (used to disambiguate buys vs sells with the same description)
    2,
    options
  );

  return searcher;
}

function findSimilarTransactions(originalDescription) {
  var limit = 3;
  if (TRANSACTION_SEARCHER === null) {
    TRANSACTION_SEARCHER = createSearchIndexWithStandardColumns({
      minTermSize: 3,
    });
  }

  const results = TRANSACTION_SEARCHER.search(originalDescription, limit);

  var previousTransactionList = [];
  results.forEach(function (result, index) {
    previousTransactionList.push({
      original_description: result.text,
      updated_description: result.updatedText,
      category: result.category,
      amount: result.amount,
    });
  });

  return previousTransactionList;
}

// Progressive batch-optimized function - reads IDs in batches of 500 until all
// target transactions are found, then writes only to contiguous row ranges.
// Optimized for large sheets where new transactions are at the top.
function writeUpdatedTransactions(transactionList, categoryList) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  var ID_BATCH_SIZE = 500;

  // --- STEP 1: Get All Column Indexes ---
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var idColIdx = headers.indexOf(TRANSACTION_ID_COL_NAME);
  var catColIdx = headers.indexOf(CATEGORY_COL_NAME);
  var descColIdx = headers.indexOf(DESCRIPTION_COL_NAME);
  var aiFlagColIdx = headers.indexOf(AI_AUTOCAT_COL_NAME);

  if (idColIdx === -1 || catColIdx === -1 || descColIdx === -1) {
    Logger.log("Error: Critical columns not found. Check your header names.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1; // Exclude header
  if (numRows < 1) return; // No data to update

  // Build set of transaction IDs we need to find
  var targetIds = {};
  for (var i = 0; i < transactionList.length; i++) {
    targetIds[transactionList[i]["transaction_id"]] = transactionList[i];
  }
  var numTargets = transactionList.length;

  // --- STEP 2: Progressive ID Loading ---
  // Read IDs in batches until we find all target transactions
  var foundRows = {}; // txId -> sheet row number (1-indexed)
  var numFound = 0;
  var rowOffset = 2; // Start after header (row 1)

  while (numFound < numTargets && rowOffset <= lastRow) {
    var batchSize = Math.min(ID_BATCH_SIZE, lastRow - rowOffset + 1);
    var idBatch = sheet.getRange(rowOffset, idColIdx + 1, batchSize, 1).getValues();

    for (var i = 0; i < idBatch.length; i++) {
      var id = idBatch[i][0];
      if (id && targetIds.hasOwnProperty(id) && !foundRows.hasOwnProperty(id)) {
        foundRows[id] = rowOffset + i; // Store actual sheet row number
        numFound++;
        if (numFound >= numTargets) break;
      }
    }

    rowOffset += batchSize;
  }

  if (numFound === 0) {
    Logger.log("No matching transactions found to update.");
    return;
  }

  Logger.log("Found " + numFound + " of " + numTargets + " transactions in first " + (rowOffset - 2) + " rows.");

  // --- STEP 3: Group Found Rows into Contiguous Ranges ---
  var rowNumbers = [];
  for (var txId in foundRows) {
    rowNumbers.push(foundRows[txId]);
  }
  rowNumbers.sort(function(a, b) { return a - b; });

  var ranges = [];
  var rangeStart = rowNumbers[0];
  var rangeEnd = rowNumbers[0];

  for (var i = 1; i < rowNumbers.length; i++) {
    if (rowNumbers[i] === rangeEnd + 1) {
      // Contiguous, extend current range
      rangeEnd = rowNumbers[i];
    } else {
      // Gap found, save current range and start new one
      ranges.push({ start: rangeStart, end: rangeEnd });
      rangeStart = rowNumbers[i];
      rangeEnd = rowNumbers[i];
    }
  }
  ranges.push({ start: rangeStart, end: rangeEnd }); // Don't forget last range

  Logger.log("Grouped into " + ranges.length + " contiguous range(s).");

  // --- STEP 4: Process Each Contiguous Range ---
  var totalUpdated = 0;

  for (var r = 0; r < ranges.length; r++) {
    var range = ranges[r];
    var rangeSize = range.end - range.start + 1;

    try {
      // Read current values for this range
      var catRange = sheet.getRange(range.start, catColIdx + 1, rangeSize, 1);
      var catValues = catRange.getValues();

      var descRange = sheet.getRange(range.start, descColIdx + 1, rangeSize, 1);
      var descValues = descRange.getValues();

      var aiFlagRange = (aiFlagColIdx !== -1) ?
        sheet.getRange(range.start, aiFlagColIdx + 1, rangeSize, 1) : null;
      var aiFlagValues = (aiFlagRange) ? aiFlagRange.getValues() : null;

      // Update values in memory
      for (var txId in foundRows) {
        var sheetRow = foundRows[txId];
        if (sheetRow >= range.start && sheetRow <= range.end) {
          var localIdx = sheetRow - range.start;
          var tx = targetIds[txId];

          // Update Category - use fallback if not in allowed list
          var newCat = tx["category"];
          if (!categoryList.includes(newCat)) {
            newCat = FALLBACK_CATEGORY;
          }
          catValues[localIdx][0] = newCat;

          // Update Description (preserve existing if not provided)
          if (tx["updated_description"]) {
            descValues[localIdx][0] = tx["updated_description"];
          }

          // Update AI Flag (if column exists)
          if (aiFlagValues) {
            aiFlagValues[localIdx][0] = "TRUE";
          }

          totalUpdated++;
        }
      }

      // Write updated values back
      catRange.setValues(catValues);
      descRange.setValues(descValues);
      if (aiFlagRange) {
        aiFlagRange.setValues(aiFlagValues);
      }

    } catch (error) {
      Logger.log("Error processing range " + range.start + "-" + range.end + ": " + error);
    }
  }

  Logger.log("Success: Updated " + totalUpdated + " transactions across " + ranges.length + " range(s).");
}

function getAllowedCategories() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var categorySheet = spreadsheet.getSheetByName(CATEGORY_SHEET_NAME);
  var headers = categorySheet.getRange("1:1").getValues()[0];

  var categoryColLetter = getColumnLetterFromColumnHeader(
    headers,
    CATEGORY_COL_NAME
  );

  var categoryListRaw = categorySheet
    .getRange(categoryColLetter + "2:" + categoryColLetter)
    .getValues();

  var categoryList = [];
  for (var i = 0; i < categoryListRaw.length; i++) {
    categoryList.push(categoryListRaw[i][0]);
  }
  return categoryList;
}

function getColumnLetterFromColumnHeader(columnHeaders, columnName) {
  var columnIndex = columnHeaders.indexOf(columnName);
  var columnLetter = "";

  let base = 26;
  let letterCharCodeBase = "A".charCodeAt(0);

  while (columnIndex >= 0) {
    columnLetter =
      String.fromCharCode((columnIndex % base) + letterCharCodeBase) +
      columnLetter;
    columnIndex = Math.floor(columnIndex / base) - 1;
  }

  return columnLetter;
}

function lookupDescAndCategoryGemini(transactionList, categoryList) {
  var transactionDict = {
    transactions: transactionList,
  };

  const request = {
    system_instruction: {
      parts: {
        text: `
        Act as an API that categorizes and cleans up bank transaction descriptions for for a personal finance app. Respond with only JSON.

        Reference the following list of allowed_categories:
        ${JSON.stringify(categoryList)}

        You will be given JSON input with a list of transaction descriptions and potentially related previously categorized transactions in the following format:
            {"transactions": [
              {
                "transaction_id": "A unique ID for this transaction"
                "original_description": "The original raw transaction description",
                "previous_transactions": "(optional) Previously cleaned up transaction descriptions and the prior 
                category used that may be related to this transaction
              }
            ]}
            For each transaction provided, follow these instructions:
            (0) If previous_transactions were provided, see if the current transaction matches a previous one closely.
                If it does, use the updated_description and category of the previous transaction exactly,
                including capitalization and punctuation.
            (1) If there is no matching previous_transaction, or none was provided suggest a better “updated_description” according to the following rules:
            (a) Use all of your knowledge and information to propose a friendly, human readable updated_description for the
              transaction given the original_description. The input often contains the name of a merchant name.
              If you know of a merchant it might be referring to, use the name of that merchant for the suggested description.
            (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous
              numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.
            (2) For each original_description, suggest a “category” for the transaction from the allowed_categories list that was provided.
            (3) If you are not confident in the suggested category after using your own knowledge and the previous transactions provided, use the cateogry "${FALLBACK_CATEGORY}"
            (4) Your response should be a JSON object and no other text.  The response object should be of the form:
            {"suggested_transactions": [
              {
                "transaction_id": "The unique ID previously provided for this transaction",
                "updated_description": "The cleaned up version of the description",
                "category": "A category selected from the allowed_categories list"
              }
            ]}
        `,
      },
    },
    contents: {
      parts: {
        text: JSON.stringify(transactionDict),
      },
    },
  };

  const jsonRequest = JSON.stringify(request);

  const options = {
    method: "POST",
    contentType: "application/json",
    payload: jsonRequest,
    muteHttpExceptions: true,
  };

  const startTime = new Date().getTime();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GOOGLE_API_KEY}`;
  var response = UrlFetchApp.fetch(url, options).getContentText();
  var parsedResponse = JSON.parse(response);

  const usage = parsedResponse.usageMetadata;
  const inputCostPerM = 0.075;
  const outputCostPerM = 0.3;
  const inputCost = (usage.promptTokenCount / 1000000) * inputCostPerM;
  const outputCost = (usage.candidatesTokenCount / 1000000) * outputCostPerM;
  const totalCost = inputCost + outputCost;
  const elapsedTime = new Date().getTime() - startTime;

  const stats = {
    elapsedTime: elapsedTime,
    numTransactions: transactionList.length,
    totalCost: totalCost,
    inputTokens: usage.promptTokenCount,
    outputTokens: usage.candidatesTokenCount,
  };

  Logger.log(stats);

  if ("error" in parsedResponse) {
    Logger.log("Error from Gemini: " + parsedResponse);
    return null;
  }

  const rawText =
    parsedResponse["candidates"][0]["content"]["parts"][0]["text"];
  const jsonStart = rawText.indexOf("{");
  const jsonEnd = rawText.lastIndexOf("}") + 1; // +1 to include the closing brace
  const cleanText = rawText.substring(jsonStart, jsonEnd);

  // Now parse the cleaned JSON
  const apiResponse = JSON.parse(cleanText);
  return apiResponse["suggested_transactions"];
}

function lookupDescAndCategoryOpenai(
  transactionList,
  categoryList,
  model = GPT_MODEL) {
  var transactionDict = {
    transactions: transactionList,
  };

  const request = {
    model: model,
    temperature: 0.2,
    top_p: 0.1,
    seed: 1,
    response_format: { type: "json_object" },
    messages: [
      {
        role: "system",
        content:
          "Act as an API that categorizes and cleans up bank transaction descriptions for for a personal finance app.",
      },
      {
        role: "system",
        content:
          "Reference the following list of allowed_categories:\n" +
          JSON.stringify(categoryList),
      },
      {
        role: "system",
        content:
          'You will be given JSON input with a list of transaction descriptions and potentially related previously categorized transactions in the following format: \
            {"transactions": [\
              {\
                "transaction_id": "A unique ID for this transaction"\
                "original_description": "The original raw transaction description",\
                "previous_transactions": "(optional) Previously cleaned up transaction descriptions and the prior \
                category used that may be related to this transaction\
              }\
            ]}\n\
            For each transaction provided, follow these instructions:\n\
            (0) If previous_transactions were provided, see if the current transaction matches a previous one closely. \
                If it does, use the updated_description and category of the previous transaction exactly, \
                including capitalization and punctuation.\
            (1) If there is no matching previous_transaction, or none was provided suggest a better “updated_description” according to the following rules:\n\
            (a) Use all of your knowledge and information to propose a friendly, human readable updated_description for the \
              transaction given the original_description. The input often contains the name of a merchant name. \
              If you know of a merchant it might be referring to, use the name of that merchant for the suggested description.\n\
            (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous \
              numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.\n\
            (2) For each original_description, suggest a “category” for the transaction from the allowed_categories list that was provided.\n\
            (3) If you are not confident in the suggested category after using your own knowledge and the previous transactions provided, use the cateogry "' +
          FALLBACK_CATEGORY +
          '"\n\n\
            (4) Your response should be a JSON object and no other text.  The response object should be of the form:\n\
            {"suggested_transactions": [\
              {\
                "transaction_id": "The unique ID previously provided for this transaction",\
                "updated_description": "The cleaned up version of the description",\
                "category": "A category selected from the allowed_categories list"\
              }\
            ]}',
      },
      {
        role: "user",
        content: JSON.stringify(transactionDict),
      },
    ],
  };

  const jsonRequest = JSON.stringify(request);

  const options = {
    method: "POST",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + OPENAI_API_KEY },
    payload: jsonRequest,
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    options
  ).getContentText();
  var parsedResponse = JSON.parse(response);

  if ("error" in parsedResponse) {
    Logger.log("Error from Open AI: " + parsedResponse["error"]["message"]);

    return null;
  } else {
    var apiResponse = JSON.parse(
      parsedResponse["choices"][0]["message"]["content"]
    );
    return apiResponse["suggested_transactions"];
  }
}
