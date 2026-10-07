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

// Generation settings, matching the Compound categorizer: deterministic output,
// as little thinking as the model allows, and a ceiling high enough that a full
// batch is never cut off mid-answer. Compound asks for "minimal" through
// OpenRouter, but Vertex rejects MINIMAL for gemini-3.8-flash (HTTP 400,
// "Thinking level is unsupported"), so LOW is the least it accepts here.
const GEMINI_TEMPERATURE = 0;
const GEMINI_THINKING_LEVEL = 'LOW';
const GEMINI_MAX_OUTPUT_TOKENS = 32000;

// Optional Jev stage. When an OpenRouter API key is set as the script property
// below, each transaction that has previous transactions is first shown to
// TypeSafe's Jev, which picks the one that is the same merchant or recurring
// payment (or none). A pick copies that transaction's category and description,
// and only the rest go to Gemini. Without the property, Gemini does everything.
// The key is read from Script Properties for the same reason as GCP_PROJECT_ID.
const OPENROUTER_API_KEY_PROPERTY = 'OPENROUTER_API_KEY';
const JEV_MODEL = 'typesafe/jev-1.13';
const JEV_URL = 'https://openrouter.ai/api/alpha/decisions';
// A pick is taken when the probability that no option matches is at most 0.5.
// Gate on that rather than on the pick's own confidence: three previous
// transactions from one merchant split the probability between them.
const JEV_TAKE_THRESHOLD = 0.5;
const JEV_CONCURRENCY = 8;
const JEV_NONE = 'none';
const JEV_NONE_OPTION =
  'No previous transaction is the same merchant or the same recurring payment as this one';
const JEV_INSTRUCTIONS =
  "Which previous transaction is the same merchant or the same recurring payment as the transaction in the state? Compare the state's bank_description with each option's bank_description: both are raw text from the bank, so the same merchant or payee recurs with the same words, minus store numbers, dates and reference codes. An option's category is how the household filed that earlier transaction; it is what the caller will copy from the option you choose, not what to match on. Prefer the option closest in amount and in day of the month when several are the same merchant. Choose none when no option is the same merchant or the same recurring payment.";

// Only transactions from this many days back are used as previous transactions.
const PRECEDENT_LOOKBACK_DAYS = 365;

// Previous transactions shown per transaction, to Jev as options and to Gemini.
// They are chosen from the top PRECEDENT_CANDIDATES search hits: when more of
// those than EXAMPLES_PER_ROW tie EXACTLY on score (within EXACT_TIE of the
// top), the closest in amount win - see closestByAmount.
const EXAMPLES_PER_ROW = 6;
const PRECEDENT_CANDIDATES = 20;
const EXACT_TIE = 0.9999;
const OPPOSITE_SIGN_DISTANCE = 1000;

// What Gemini answers when it declines to pick a category. It is never on the
// allowed list, so writeUpdatedTransactions files it under FALLBACK_CATEGORY.
const DECLINED_CATEGORY = "to-be-categorized";

// Sheet Names
const TRANSACTION_SHEET_NAME = "Transactions";
const CATEGORY_SHEET_NAME = "Categories";

// Column Names
const TRANSACTION_ID_COL_NAME = "Transaction ID";
const ORIGINAL_DESCRIPTION_COL_NAME = "Full Description";
const DESCRIPTION_COL_NAME = "Description";
const CATEGORY_COL_NAME = "Category";
const GROUP_COL_NAME = "Group";
const AI_AUTOCAT_COL_NAME = "AI AutoCat";
const DATE_COL_NAME = "Date";
const AMOUNT_COL_NAME = "Amount";

// Fallback Transaction Category (to be used when we don't know how to categorize a transaction)
const FALLBACK_CATEGORY = "To Be Categorized";

// Other Misc Paramaters
const MAX_BATCH_SIZE = 50;
var TRANSACTION_SEARCHER = null;

function categorizeUncategorizedTransactions() {
  assignMissingTransactionIds();
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
    var txn = uncategorizedTransactions[i];
    var entry = {
      transaction_id: txn.transaction_id,
      original_description: txn.original_description,
    };
    if (txn.amount !== undefined) entry.amount = txn.amount;
    if (txn.date) entry.date = txn.date;
    entry.previous_transactions = findSimilarTransactions(
      txn.original_description,
      txn.amount
    );
    transactionList.push(entry);
  }

  Logger.log(
    "Processing this set of transactions and similar transactions:"
  );
  Logger.log(transactionList);

  var categoryList = getAllowedCategories();

  // Stage 1 (optional): Jev settles the transactions that match a previous one.
  var byPrecedent = {
    suggestions: [],
    decisions: [],
    asked: 0,
    taken: 0,
    failed: 0,
    cost: 0,
  };
  var openRouterKey = PropertiesService.getScriptProperties().getProperty(
    OPENROUTER_API_KEY_PROPERTY
  );
  if (openRouterKey) {
    Logger.log("Asking Jev (" + JEV_MODEL + ") which previous transaction matches...");
    byPrecedent = askPrecedents(transactionList, categoryList, openRouterKey);
    Logger.log({
      jevAsked: byPrecedent.asked,
      jevTaken: byPrecedent.taken,
      jevFailed: byPrecedent.failed,
      jevCost: byPrecedent.cost,
    });
    if (byPrecedent.decisions.length > 0) {
      // One line per transaction asked: what Jev chose, the probability that
      // some option matches (1 - P(none), the number the take threshold is
      // checked against), Jev's own confidence, and whether it was taken.
      Logger.log("Jev decisions:");
      Logger.log(byPrecedent.decisions);
    }
    if (byPrecedent.suggestions.length > 0) {
      Logger.log("Jev matched these to a previous transaction:");
      Logger.log(byPrecedent.suggestions);
    }
  }

  // Write Jev's answers now, so they are saved even if the Gemini call fails
  // or the run is cut off.
  if (byPrecedent.suggestions.length > 0) {
    Logger.log("Writing Jev's matches into your sheet...");
    writeUpdatedTransactions(byPrecedent.suggestions, categoryList);
  }

  // Stage 2: Gemini answers for everything Jev did not settle.
  var taken = Object.create(null);
  byPrecedent.suggestions.forEach(function (s) {
    taken[s.transaction_id] = true;
  });
  var remaining = transactionList.filter(function (t) {
    return !taken[t.transaction_id];
  });

  var byGemini = [];
  if (remaining.length > 0) {
    Logger.log(
      "Using Gemini (" + GEMINI_MODEL + ") on Vertex AI for " +
        remaining.length + " transaction(s)"
    );
    byGemini = lookupDescAndCategoryGemini(
      remaining,
      getPromptCategories(categoryList)
    );
    if (byGemini == null) {
      // Jev's matches are already written; the rest stay uncategorized for the
      // next run.
      byGemini = [];
    } else {
      Logger.log(
        "Gemini returned the following sugested categories and descriptions:"
      );
      Logger.log(byGemini);
    }
  }

  if (byGemini.length > 0) {
    Logger.log("Writing Gemini's answers into your sheet...");
    writeUpdatedTransactions(byGemini, categoryList);
  }

  if (byPrecedent.suggestions.length > 0 || byGemini.length > 0) {
    Logger.log("Finished updating your sheet!");
  }
}

/**
 * Formats a sheet or gviz date value as YYYY-MM-DD. Sheet reads give Date
 * objects (in the script's time zone); gviz gives strings like
 * "Date(2026,7,14)" with a zero-based month. Anything else that is non-empty
 * is passed through as text; empty values give "".
 */
function isoDate(value) {
  if (value === null || value === undefined || value === "") return "";
  var pad = function (n) {
    return (n < 10 ? "0" : "") + n;
  };
  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "";
    return (
      value.getFullYear() + "-" + pad(value.getMonth() + 1) + "-" + pad(value.getDate())
    );
  }
  var m = /^Date\((\d+),(\d+),(\d+)/.exec(String(value));
  if (m) {
    return m[1] + "-" + pad(Number(m[2]) + 1) + "-" + pad(Number(m[3]));
  }
  return String(value);
}

// Transaction IDs this script assigns, to rows it is about to send that have
// none (usually rows added by hand). Tiller's own hand-entered rows use
// "manual:<uuid>"; this prefix marks where these came from and cannot collide.
const ASSIGNED_ID_PREFIX = "autocat:";

const CROCKFORD = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";

// A ULID: 10 characters of millisecond timestamp, then 16 of randomness, in
// Crockford base 32. Sorts by creation time.
function ulid(now) {
  var time = now === undefined ? Date.now() : now;
  var out = "";
  for (var i = 0; i < 10; i++) {
    out = CROCKFORD.charAt(time % 32) + out;
    time = Math.floor(time / 32);
  }
  for (var j = 0; j < 16; j++) {
    out += CROCKFORD.charAt(Math.floor(Math.random() * 32));
  }
  return out;
}

/**
 * Gives a Transaction ID to every row this run is about to send that has none,
 * so that its answer can be written back by id like any other row.
 *
 * "About to send" follows the same rule as getTransactionsToCategorize: rows
 * with a Full Description and no Category, in sheet order, up to
 * MAX_BATCH_SIZE of them. Only the empty Transaction ID cells of those rows are
 * written (invariant (b)), and the writes are flushed so the gviz query that
 * follows sees them.
 *
 * @returns {number} how many ids were assigned
 */
function assignMissingTransactionIds() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    TRANSACTION_SHEET_NAME
  );
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idColIdx = headers.indexOf(TRANSACTION_ID_COL_NAME);
  var fullDescColIdx = headers.indexOf(ORIGINAL_DESCRIPTION_COL_NAME);
  var catColIdx = headers.indexOf(CATEGORY_COL_NAME);
  if (idColIdx === -1 || fullDescColIdx === -1 || catColIdx === -1) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var blank = function (v) {
    return v === null || v === undefined || v === "";
  };
  var BATCH = 2000;
  var cells = [];
  var seen = 0;

  for (var start = 2; start <= lastRow && seen < MAX_BATCH_SIZE; start += BATCH) {
    var size = Math.min(BATCH, lastRow - start + 1);
    var ids = sheet.getRange(start, idColIdx + 1, size, 1).getValues();
    var descs = sheet.getRange(start, fullDescColIdx + 1, size, 1).getValues();
    var cats = sheet.getRange(start, catColIdx + 1, size, 1).getValues();

    for (var r = 0; r < size && seen < MAX_BATCH_SIZE; r++) {
      if (blank(descs[r][0]) || !blank(cats[r][0])) continue;
      seen++;
      if (blank(ids[r][0])) {
        cells.push({
          row: start + r,
          column: idColIdx + 1,
          value: ASSIGNED_ID_PREFIX + ulid(),
        });
      }
    }
  }

  if (cells.length === 0) return 0;

  planCellWrites(cells).forEach(function (block) {
    sheet
      .getRange(block.row, block.column, block.values.length, block.values[0].length)
      .setValues(block.values);
  });
  SpreadsheetApp.flush();

  Logger.log(
    "Assigned a Transaction ID to " + cells.length + " row(s) that had none: rows " +
      cells.map(function (c) { return c.row; }).join(", ")
  );
  return cells.length;
}

// Gets up to MAX_BATCH_SIZE transactions that have an original description but
// no category set, as {transaction_id, original_description, amount?, date}.
// Amount and Date are optional columns, read only when present.
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

  // Invariant (a): optional columns resolve to "" and are left out of the query.
  var amountColLetter = getColumnLetterFromColumnHeader(headers, AMOUNT_COL_NAME);
  var dateColLetter = getColumnLetterFromColumnHeader(headers, DATE_COL_NAME);
  var selected = [txnIDColLetter, origDescColLetter];
  var amountAt = amountColLetter ? selected.push(amountColLetter) - 1 : -1;
  var dateAt = dateColLetter ? selected.push(dateColLetter) - 1 : -1;

  var queryString =
    "SELECT " +
    selected.join(", ") +
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

  return uncategorizedTransactions.map(function (row) {
    var txn = { transaction_id: row[0], original_description: row[1] };
    if (amountAt !== -1 && typeof row[amountAt] === "number") {
      txn.amount = row[amountAt];
    }
    txn.date = dateAt !== -1 ? isoDate(row[dateAt]) : "";
    return txn;
  });
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
    Object.assign(
      { since: lookbackStart(PRECEDENT_LOOKBACK_DAYS) },
      options || {}
    )
  );

  return searcher;
}

/**
 * Distance between two amounts for choosing previous transactions: how far
 * apart they are on a log scale, or OPPOSITE_SIGN_DISTANCE when one is money
 * in and the other money out.
 */
function amountDistance(a, b) {
  return Math.sign(a) === Math.sign(b)
    ? Math.abs(Math.log((Math.abs(a) + 1) / (Math.abs(b) + 1)))
    : OPPOSITE_SIGN_DISTANCE;
}

/**
 * Chooses up to `slots` previous transactions from TF-IDF hits ranked best
 * first. Usually that is just the top `slots`. But when more hits than that tie
 * exactly on score - every "CHECK #1234" looks the same to word overlap - the
 * tied ones are re-ordered by how close their amount is, the closest of each
 * category is taken first, and any free slots are filled from the same order.
 * Mirrors closestByAmount in Compound's categorizePrepare.ts.
 */
function closestByAmount(hits, amount, slots) {
  if (hits.length === 0) return [];
  var top = hits[0].score;
  var tied = hits.filter(function (hit) {
    return hit.score >= top * EXACT_TIE;
  });
  if (tied.length <= slots) return hits.slice(0, slots);

  var amountOf = function (hit) {
    return typeof hit.amount === "number" ? hit.amount : 0;
  };
  // Array.prototype.sort is stable in V8, so equal distances keep rank order.
  var ranked = tied.slice().sort(function (a, b) {
    return amountDistance(amountOf(a), amount) - amountDistance(amountOf(b), amount);
  });

  var chosen = [];
  var categories = Object.create(null);
  ranked.forEach(function (hit) {
    var category = hit.category || "";
    if (chosen.length < slots && !categories[category]) {
      categories[category] = true;
      chosen.push(hit);
    }
  });
  ranked.forEach(function (hit) {
    if (chosen.length < slots && chosen.indexOf(hit) === -1) chosen.push(hit);
  });
  return ranked.filter(function (hit) {
    return chosen.indexOf(hit) !== -1;
  });
}

function findSimilarTransactions(originalDescription, amount) {
  if (TRANSACTION_SEARCHER === null) {
    TRANSACTION_SEARCHER = createSearchIndexWithStandardColumns({
      minTermSize: 3,
    });
  }

  const results = closestByAmount(
    TRANSACTION_SEARCHER.search(originalDescription, PRECEDENT_CANDIDATES),
    typeof amount === "number" ? amount : 0,
    EXAMPLES_PER_ROW
  );

  var previousTransactionList = [];
  results.forEach(function (result, index) {
    previousTransactionList.push({
      original_description: result.text,
      updated_description: result.updatedText,
      category: result.category,
      amount: result.amount,
      date: isoDate(result.date),
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

// The category list as the model sees it: [{name, group}], in sheet order. The
// group comes from the Categories sheet's Group column and is left off when
// that column is absent or the cell is blank. Names, not ids, are what the
// model reads and answers with.
function getPromptCategories(categoryList) {
  var categorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CATEGORY_SHEET_NAME
  );
  var groupOf = Object.create(null);

  if (categorySheet) {
    var headers = categorySheet.getRange("1:1").getValues()[0];
    var categoryColLetter = getColumnLetterFromColumnHeader(headers, CATEGORY_COL_NAME);
    var groupColLetter = getColumnLetterFromColumnHeader(headers, GROUP_COL_NAME);

    if (categoryColLetter && groupColLetter) {
      var names = categorySheet
        .getRange(categoryColLetter + "2:" + categoryColLetter)
        .getValues();
      var groups = categorySheet
        .getRange(groupColLetter + "2:" + groupColLetter)
        .getValues();
      for (var i = 0; i < names.length; i++) {
        var name = names[i][0];
        var group = groups[i] && groups[i][0];
        if (name !== "" && name !== null && group !== "" && group != null) {
          groupOf[name] = String(group);
        }
      }
    }
  }

  return categoryList.map(function (name) {
    var entry = { name: name };
    if (groupOf[name]) entry.group = groupOf[name];
    return entry;
  });
}

// The description to copy from a previous transaction: its cleaned-up
// Description, or the raw text when it has none.
function displayDescription(previous) {
  return previous.updated_description || previous.original_description || "";
}

/**
 * Builds the Jev request for one transaction: the transaction as the state, its
 * previous transactions as options p1..pN, plus "none". Jev sees raw bank text
 * on both sides and the category by name; the cleaned description is not shown,
 * it is what gets copied from the option Jev picks.
 */
function precedentRequest(transaction) {
  var state = {
    transaction: { bank_description: transaction.original_description },
  };
  if (transaction.amount !== undefined) state.transaction.amount = transaction.amount;
  if (transaction.date) state.transaction.date = transaction.date;

  var criteria = {};
  transaction.previous_transactions.forEach(function (p, i) {
    criteria["p" + (i + 1)] = {
      bank_description: p.original_description,
      category: p.category,
      amount: p.amount,
      date: p.date,
    };
  });
  criteria[JEV_NONE] = JEV_NONE_OPTION;

  return {
    model: JEV_MODEL,
    state: state,
    questions: {
      precedent: {
        type: "choice",
        instructions: JEV_INSTRUCTIONS,
        criteria: criteria,
      },
    },
  };
}

/**
 * Reads Jev's answer and returns the previous transaction it picked, or null
 * when it chose none, picked something unrecognised, or was not sure enough.
 */
function takenPrecedent(answer, previous) {
  if (!answer || typeof answer.choice !== "string") return null;
  if (answer.choice === JEV_NONE) return null;
  var m = /^p(\d+)$/.exec(answer.choice);
  if (!m) return null;
  var picked = previous[Number(m[1]) - 1];
  if (!picked) return null;
  var none = answer.probabilities && answer.probabilities[JEV_NONE];
  var matchProbability = 1 - (typeof none === "number" ? none : 0);
  return matchProbability >= JEV_TAKE_THRESHOLD ? picked : null;
}

/**
 * A one-line summary of Jev's answer for the log: the raw description, the
 * option it chose (with that option's category), the match probability
 * 1 - P(none), Jev's confidence, and whether the pick was taken.
 */
function jevDecision(transaction, answer, taken) {
  var round = function (n) {
    return typeof n === "number" ? Math.round(n * 1000) / 1000 : null;
  };
  var choice = answer && typeof answer.choice === "string" ? answer.choice : "";
  var m = /^p(\d+)$/.exec(choice);
  var option = m ? transaction.previous_transactions[Number(m[1]) - 1] : null;
  var none = answer && answer.probabilities && answer.probabilities[JEV_NONE];
  return {
    transaction_id: transaction.transaction_id,
    description: transaction.original_description,
    choice: choice + (option ? " (" + option.category + ")" : ""),
    match: round(1 - (typeof none === "number" ? none : 0)),
    confidence: round(answer && answer.confidence),
    taken: taken,
  };
}

/**
 * Stage 1: asks Jev, through OpenRouter, which previous transaction each
 * transaction matches. Only transactions with previous transactions are asked,
 * JEV_CONCURRENCY at a time. A pick is taken only when its category is still on
 * the allowed list and is not FALLBACK_CATEGORY - copying a "To Be Categorized"
 * would just repeat an earlier non-answer, so those go on to Gemini.
 *
 * A failed request is counted and that transaction falls through to Gemini.
 *
 * @returns {{suggestions: Array, asked: number, taken: number, failed: number, cost: number}}
 */
function askPrecedents(transactionList, categoryList, apiKey) {
  var candidates = transactionList.filter(function (t) {
    return t.previous_transactions && t.previous_transactions.length > 0;
  });
  var suggestions = [];
  var decisions = [];
  var failed = 0;
  var cost = 0;

  for (var at = 0; at < candidates.length; at += JEV_CONCURRENCY) {
    var chunk = candidates.slice(at, at + JEV_CONCURRENCY);
    var responses = UrlFetchApp.fetchAll(
      chunk.map(function (t) {
        return {
          url: JEV_URL,
          method: "post",
          contentType: "application/json",
          headers: { Authorization: "Bearer " + apiKey },
          payload: JSON.stringify(precedentRequest(t)),
          muteHttpExceptions: true,
        };
      })
    );

    for (var i = 0; i < chunk.length; i++) {
      var json = null;
      try {
        json = JSON.parse(responses[i].getContentText());
      } catch (e) {}

      if (responses[i].getResponseCode() != 200 || !json || json.error) {
        failed++;
        Logger.log(
          "Jev request failed for " + chunk[i].transaction_id + " (HTTP " +
            responses[i].getResponseCode() + "): " +
            String(responses[i].getContentText()).slice(0, 300)
        );
        continue;
      }

      if (json.usage && typeof json.usage.cost === "number") cost += json.usage.cost;

      var answer = json.answers && json.answers.precedent;
      var picked = takenPrecedent(answer, chunk[i].previous_transactions);
      var take =
        !!picked &&
        picked.category !== FALLBACK_CATEGORY &&
        categoryList.includes(picked.category);
      decisions.push(jevDecision(chunk[i], answer, take));
      if (take) {
        suggestions.push({
          transaction_id: chunk[i].transaction_id,
          updated_description: displayDescription(picked),
          category: picked.category,
        });
      }
    }
  }

  return {
    suggestions: suggestions,
    decisions: decisions,
    asked: candidates.length,
    taken: suggestions.length,
    failed: failed,
    cost: cost,
  };
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

// The categorizer prompt, kept in step with Compound's categorize prompt
// (compound/primitive/alpha/prompts/categorize.toml). Differences: categories
// are named rather than given ids, and the Plaid transaction-kind hint is left
// out because a Tiller sheet has no such column.
const GEMINI_SYSTEM_PROMPT = `Act as an API that cleans up bank transaction descriptions and files them into
a household's own categories. Respond with ONLY JSON.

The input JSON has this shape:
{"allowed_categories": [{"name": "...", "group": "..."}],
 "transactions": [
   {"transaction_id": "...",
    "original_description": "the raw bank description",
    "amount": 12.34,
    "date": "2026-08-14",
    "previous_transactions": [
      {"bank_description": "the raw bank description",
       "description": "...", "category": "Groceries",
       "amount": 1.23, "date": "2026-07-14"}
    ]}
 ]}

original_description is raw text straight from the bank. Each previous
transaction carries its own raw text as bank_description, so match
original_description against bank_description, raw against raw: the same
merchant or payee recurs with the same words, minus store numbers, dates and
reference codes. A previous transaction's description is NOT raw: it is the
cleaned-up name this household already uses for that transaction, and it is
the wording to reuse. Its category is the name of the household category it was
filed in, as it appears in allowed_categories.

amount is signed from the account holder's point of view: negative means money
left the account, positive means money arrived.

previous_transactions are this household's own earlier decisions, found by word
overlap with the raw description. They are the strongest signal available,
because they are how this household has chosen to name and file this kind of
transaction.

For each transaction, answer with both an updated_description and a category.

Choosing updated_description:
(1) If a previous transaction's bank_description plausibly is the same merchant
    or the same recurring payment, reuse its description EXACTLY, including
    capitalization and punctuation. Consistency matters more than your own
    phrasing.
(2) Otherwise write a friendly, human-readable name for the transaction. The
    raw description usually contains a merchant; if you recognize it, use that
    merchant's proper name.
(3) Keep it as simple as possible. Remove punctuation, extraneous numbers,
    location information, abbreviations like "Inc." or "LLC", store and
    reference numbers, and account numbers.
(4) If the raw description tells you nothing you can clean up, return it
    unchanged rather than inventing a merchant.

Choosing category, in this order:
(1) If the previous transactions that plausibly are the same merchant or the
    same recurring payment agree on a category, use that category.
(2) If they disagree, prefer the one closest in amount and in day of the month.
    Recurring bills land at the same point in the month for similar amounts.
(3) If no previous transaction convinces you, fall back on general knowledge of
    the merchant, using the sign of the amount as a hint.
(4) If you are still unsure, answer "${DECLINED_CATEGORY}" for that transaction.
    Leaving it for a person is a better outcome than a wrong category.

Answering "${DECLINED_CATEGORY}" for the category does not excuse you from
cleaning the description: give your best updated_description either way.

Answer for EVERY transaction you were given, in the order you were given them.
"${DECLINED_CATEGORY}" is how you decline; a missing entry is not, and is read as
an answer that went astray.

Every other category you return must be the "name" of an entry in
allowed_categories, spelled exactly as given. Never invent one.

Respond with a JSON object of exactly this form and no other text:
{"suggested_transactions": [
  {"transaction_id": "...", "updated_description": "...", "category": "..."}
]}
`;

// Vertex AI's structured-output schema for the answer above.
const GEMINI_RESPONSE_SCHEMA = {
  type: "OBJECT",
  required: ["suggested_transactions"],
  properties: {
    suggested_transactions: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        required: ["transaction_id", "updated_description", "category"],
        propertyOrdering: ["transaction_id", "updated_description", "category"],
        properties: {
          transaction_id: { type: "STRING" },
          updated_description: { type: "STRING" },
          category: { type: "STRING" },
        },
      },
    },
  },
};

/**
 * The user message for Gemini: the allowed categories and, per transaction,
 * its raw description, amount and date, plus its previous transactions as
 * {bank_description (raw), description (cleaned), category (name), amount, date}.
 */
function geminiPayload(transactionList, promptCategories) {
  return {
    allowed_categories: promptCategories,
    transactions: transactionList.map(function (t) {
      var entry = {
        transaction_id: t.transaction_id,
        original_description: t.original_description,
      };
      if (t.amount !== undefined) entry.amount = t.amount;
      if (t.date) entry.date = t.date;
      entry.previous_transactions = (t.previous_transactions || []).map(
        function (p) {
          return {
            bank_description: p.original_description,
            description: displayDescription(p),
            category: p.category,
            amount: p.amount,
            date: p.date,
          };
        }
      );
      return entry;
    }),
  };
}

function lookupDescAndCategoryGemini(transactionList, promptCategories) {
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

  const request = {
    systemInstruction: {
      parts: [{ text: GEMINI_SYSTEM_PROMPT }],
    },
    contents: [
      {
        role: "user",
        parts: [
          { text: JSON.stringify(geminiPayload(transactionList, promptCategories)) },
        ],
      },
    ],
    generationConfig: {
      temperature: GEMINI_TEMPERATURE,
      maxOutputTokens: GEMINI_MAX_OUTPUT_TOKENS,
      responseMimeType: "application/json",
      responseSchema: GEMINI_RESPONSE_SCHEMA,
      thinkingConfig: { thinkingLevel: GEMINI_THINKING_LEVEL },
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
  // Thought summaries, if the model ever returns them, are marked `thought`;
  // the answer is in the other parts.
  const rawText = parts
    .filter(function (part) {
      return !part.thought && typeof part.text === "string";
    })
    .map(function (part) {
      return part.text;
    })
    .join("");
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
