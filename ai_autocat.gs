// Two invariants this script relies on. Please preserve them when editing:
//
// (a) Columns can appear in any order. Never assume a column position - always
//     resolve a column from its header name via getColumnLetterFromColumnHeader,
//     and treat an empty result as "this column is not present".
//
// (b) Only ever write the specific cells you intend to change. Never write a
//     whole row, even one padded with blanks: Tiller sheets commonly define
//     ARRAYFORMULA at the column level, and writing a literal into such a column
//     permanently breaks the formula for the entire sheet.

// Google Cloud / Vertex AI Settings
// There is no API key here on purpose. Requests are authenticated with Application
// Default Credentials - ScriptApp.getOAuthToken() returns a token for whoever runs
// the script, and Vertex AI authorizes it via IAM. See the README for setup.
// Your Google Cloud project ID is read from Script Properties rather than being
// hardcoded here, so it never ends up in source control. Set it once in the Apps
// Script editor: Project Settings -> Script Properties -> Add script property,
// with the name GCP_PROJECT_ID. See the README.
const GCP_PROJECT_ID_PROPERTY = 'GCP_PROJECT_ID';

// 'global' routes to whichever region has capacity - best availability and the
// widest model support. Use a specific region (e.g. 'us-west1') if you need your
// requests to stay in one geography.
const GCP_LOCATION = 'global';
const GEMINI_MODEL = 'gemini-3.8-flash'; // Can be any model Vertex AI publishes

// Pricing for GEMINI_MODEL on Vertex AI, in US dollars per million tokens.
// Used only for the cost estimate written to the log. Check current rates at
// https://cloud.google.com/vertex-ai/generative-ai/pricing
//
// These are gemini-3.8-flash introductory rates, which run through 2026-12-31.
// On 2027-01-01 they go to 1.5 and 7.5 - until these are updated the logged
// estimate will read half of what you are actually billed.
const INPUT_COST_PER_M_TOKENS = 0.75;
const OUTPUT_COST_PER_M_TOKENS = 3.75;

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

  Logger.log("Using Gemini (" + GEMINI_MODEL + ") on Vertex AI");

  var updatedTransactions = lookupDescAndCategoryGemini(
    transactionList,
    categoryList
  );

  if (updatedTransactions != null) {
    Logger.log(
      "Gemini returned the following sugested categories and descriptions:"
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

/**
 * Groups individual cell writes into the smallest set of rectangular blocks that
 * covers exactly those cells and no others.
 *
 * This exists to serve invariant (b). Batching is safe only when every cell in
 * the block is one we meant to write: a block that spanned an untouched column
 * or row would overwrite it, and on a Tiller sheet that usually means destroying
 * an ARRAYFORMULA. Cells are merged horizontally into runs of adjacent columns,
 * then those runs are merged vertically when consecutive rows share the same
 * span. Anything with a gap stays a separate write.
 *
 * @param {Array<{row: number, column: number, value: *}>} cells 1-based cells.
 * @returns {Array<{row: number, column: number, values: Array<Array<*>>}>}
 */
function planCellWrites(cells) {
  if (cells.length === 0) {
    return [];
  }

  // Group by row, then split each row into runs of adjacent columns.
  var byRow = {};
  for (var i = 0; i < cells.length; i++) {
    var cell = cells[i];
    if (!byRow[cell.row]) {
      byRow[cell.row] = [];
    }
    byRow[cell.row].push(cell);
  }

  var runs = [];
  var rowNumbers = Object.keys(byRow)
    .map(Number)
    .sort(function (a, b) {
      return a - b;
    });

  for (var r = 0; r < rowNumbers.length; r++) {
    var row = rowNumbers[r];
    var rowCells = byRow[row].sort(function (a, b) {
      return a.column - b.column;
    });

    var current = null;
    for (var c = 0; c < rowCells.length; c++) {
      if (current && rowCells[c].column === current.column + current.values.length) {
        current.values.push(rowCells[c].value);
      } else {
        current = {
          row: row,
          column: rowCells[c].column,
          values: [rowCells[c].value],
        };
        runs.push(current);
      }
    }
  }

  // Merge runs downward when the row below covers exactly the same columns.
  var blocks = [];
  var consumed = {};

  for (var j = 0; j < runs.length; j++) {
    if (consumed[j]) {
      continue;
    }

    var block = {
      row: runs[j].row,
      column: runs[j].column,
      values: [runs[j].values],
    };

    var nextRow = runs[j].row + 1;
    for (var k = j + 1; k < runs.length; k++) {
      if (consumed[k] || runs[k].row !== nextRow) {
        continue;
      }
      if (
        runs[k].column !== block.column ||
        runs[k].values.length !== block.values[0].length
      ) {
        continue;
      }
      block.values.push(runs[k].values);
      consumed[k] = true;
      nextRow++;
    }

    blocks.push(block);
  }

  return blocks;
}

// Reads the transaction ID column in batches until every target transaction has
// been located - new transactions sit at the top of a Tiller sheet, so this
// usually touches only the first batch - then writes just the cells that change.
function writeUpdatedTransactions(transactionList, categoryList) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    TRANSACTION_SHEET_NAME
  );
  var ID_BATCH_SIZE = 500;

  // --- STEP 1: Resolve column positions ---
  // Invariant (a): every column is resolved from its header, never assumed.
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var idColIdx = headers.indexOf(TRANSACTION_ID_COL_NAME);
  var catColIdx = headers.indexOf(CATEGORY_COL_NAME);
  var descColIdx = headers.indexOf(DESCRIPTION_COL_NAME);
  // Optional column: indexOf returns -1 when it is absent.
  var aiFlagColIdx = headers.indexOf(AI_AUTOCAT_COL_NAME);

  if (idColIdx === -1 || catColIdx === -1 || descColIdx === -1) {
    Logger.log("Error: Critical columns not found. Check your header names.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return; // No data rows
  }

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
    var idBatch = sheet
      .getRange(rowOffset, idColIdx + 1, batchSize, 1)
      .getValues();

    for (var b = 0; b < idBatch.length; b++) {
      var id = idBatch[b][0];
      // hasOwnProperty via Object.prototype, so that an id colliding with an
      // inherited member name cannot produce a false match.
      if (
        id !== "" &&
        Object.prototype.hasOwnProperty.call(targetIds, id) &&
        !Object.prototype.hasOwnProperty.call(foundRows, id)
      ) {
        foundRows[id] = rowOffset + b; // Store actual sheet row number
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

  Logger.log(
    "Found " +
      numFound +
      " of " +
      numTargets +
      " transactions in first " +
      (rowOffset - 2) +
      " rows."
  );

  // --- STEP 3: Collect the cells we intend to change ---
  // Invariant (b): nothing outside this list may be written. planCellWrites
  // batches these into rectangles only where every cell in the rectangle is one
  // of them, so a column or row we do not own is never overwritten - which on a
  // Tiller sheet would mean destroying an ARRAYFORMULA.
  var cellsToWrite = [];

  for (var txId in foundRows) {
    var transactionRow = foundRows[txId];
    var tx = targetIds[txId];

    var updatedCategory = tx["category"];
    if (!categoryList.includes(updatedCategory)) {
      updatedCategory = FALLBACK_CATEGORY;
    }

    cellsToWrite.push({
      row: transactionRow,
      column: catColIdx + 1,
      value: updatedCategory,
    });

    // Leave the existing description untouched when the model did not return
    // one, rather than blanking it or rewriting it with its current value.
    if (tx["updated_description"]) {
      cellsToWrite.push({
        row: transactionRow,
        column: descColIdx + 1,
        value: tx["updated_description"],
      });
    }

    if (aiFlagColIdx !== -1) {
      cellsToWrite.push({
        row: transactionRow,
        column: aiFlagColIdx + 1,
        value: "TRUE",
      });
    }
  }

  // --- STEP 4: Write ---
  var blocks = planCellWrites(cellsToWrite);

  for (var w = 0; w < blocks.length; w++) {
    try {
      sheet
        .getRange(
          blocks[w].row,
          blocks[w].column,
          blocks[w].values.length,
          blocks[w].values[0].length
        )
        .setValues(blocks[w].values);
    } catch (error) {
      Logger.log(
        "Error writing block at row " +
          blocks[w].row +
          ", column " +
          blocks[w].column +
          ": " +
          error
      );
    }
  }

  Logger.log(
    "Success: Updated " +
      numFound +
      " transactions in " +
      blocks.length +
      " write(s)."
  );
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

  // The open-ended range above runs to the bottom of the sheet, so most of what
  // comes back is blank. Drop those rather than padding the model prompt with
  // hundreds of empty strings on every call.
  var categoryList = [];
  for (var i = 0; i < categoryListRaw.length; i++) {
    var category = categoryListRaw[i][0];
    if (category !== "" && category !== null) {
      categoryList.push(category);
    }
  }
  return categoryList;
}

// Resolves a column header name to its A1 column letter, supporting invariant (a).
// Returns "" when the column is not present, so callers testing for an optional
// column must check truthiness rather than comparing against null.
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
  const projectId = PropertiesService.getScriptProperties().getProperty(
    GCP_PROJECT_ID_PROPERTY
  );

  if (!projectId) {
    Logger.log(
      "The " +
        GCP_PROJECT_ID_PROPERTY +
        " script property is not set. In the Apps Script editor go to " +
        "Project Settings -> Script Properties and add it, using your Google " +
        "Cloud project ID as the value. See the README for full setup steps."
    );
    return null;
  }

  var transactionDict = {
    transactions: transactionList,
  };

  const request = {
    systemInstruction: {
      parts: [
        {
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
      ],
    },
    contents: [
      {
        role: "user",
        parts: [{ text: JSON.stringify(transactionDict) }],
      },
    ],
    generationConfig: {
      responseMimeType: "application/json",
    },
  };

  const options = {
    method: "POST",
    contentType: "application/json",
    // Application Default Credentials: this token belongs to whoever is running
    // the script, and Vertex AI checks their IAM role on GCP_PROJECT_ID.
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(request),
    muteHttpExceptions: true,
  };

  // Regional endpoints are prefixed with the region; the 'global' endpoint is not.
  const host =
    GCP_LOCATION == "global"
      ? "aiplatform.googleapis.com"
      : GCP_LOCATION + "-aiplatform.googleapis.com";

  const url =
    "https://" +
    host +
    "/v1/projects/" +
    projectId +
    "/locations/" +
    GCP_LOCATION +
    "/publishers/google/models/" +
    GEMINI_MODEL +
    ":generateContent";

  const startTime = new Date().getTime();
  const response = UrlFetchApp.fetch(url, options);
  const elapsedTime = new Date().getTime() - startTime;

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode != 200) {
    Logger.log(
      "Error from Vertex AI (HTTP " + responseCode + "): " + responseText
    );
    return null;
  }

  const parsedResponse = JSON.parse(responseText);
  if ("error" in parsedResponse) {
    Logger.log("Error from Vertex AI: " + JSON.stringify(parsedResponse.error));
    return null;
  }

  logUsageStats(parsedResponse.usageMetadata, transactionList.length, elapsedTime);

  const candidate = parsedResponse.candidates && parsedResponse.candidates[0];
  const parts = candidate && candidate.content && candidate.content.parts;
  if (!parts || parts.length == 0) {
    Logger.log(
      "Vertex AI returned no usable content. Full response: " + responseText
    );
    return null;
  }

  // responseMimeType asks for bare JSON, but trim anything outside the outermost
  // braces in case the model still wraps it in prose or a code fence.
  const rawText = parts[0].text;
  const jsonStart = rawText.indexOf("{");
  const jsonEnd = rawText.lastIndexOf("}") + 1; // +1 to include the closing brace
  const cleanText = rawText.substring(jsonStart, jsonEnd);

  const apiResponse = JSON.parse(cleanText);
  return apiResponse["suggested_transactions"];
}

function logUsageStats(usage, numTransactions, elapsedTime) {
  if (!usage) {
    return;
  }

  // Gemini bills thinking tokens at the output rate, and reports them separately
  // from candidatesTokenCount. The 3.x models think more than 2.5 did, so this
  // term is a larger share of the cost than it used to be.
  const inputTokens = usage.promptTokenCount || 0;
  const outputTokens =
    (usage.candidatesTokenCount || 0) + (usage.thoughtsTokenCount || 0);

  const inputCost = (inputTokens / 1000000) * INPUT_COST_PER_M_TOKENS;
  const outputCost = (outputTokens / 1000000) * OUTPUT_COST_PER_M_TOKENS;

  const stats = {
    elapsedTime: elapsedTime,
    numTransactions: numTransactions,
    totalCost: inputCost + outputCost,
    inputTokens: inputTokens,
    outputTokens: outputTokens,
  };

  Logger.log(stats);
}
