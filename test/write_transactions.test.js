/**
 * Tests for writeUpdatedTransactions and the cell-write planner.
 *
 * The invariants under test:
 *   (a) Columns can appear in any order and are resolved by header name.
 *   (b) A write must never touch a cell the script did not intend to write.
 *       Tiller sheets frequently drive whole columns from a single
 *       ARRAYFORMULA, and writing a literal into one destroys it.
 */
const test = require("node:test");
const assert = require("node:assert");
const {
  FakeSheet,
  FakeSpreadsheet,
  loadScripts,
  readWorkingTree,
  readAtRevision,
  columnIndexToLetter,
} = require("./harness");

// The implementation these tests characterise, i.e. the last commit before the
// Vertex AI / performance work. Pinned to a SHA rather than HEAD so the
// comparison stays meaningful once these changes are themselves committed.
const BASELINE_REV = process.env.BASELINE_REV || "9ac70da1025494e53d9bdd16595c400cf87dd4bf";

const CATEGORIES = [
  "Groceries",
  "Restaurants",
  "Utilities",
  "Shopping",
  "Travel",
  "To Be Categorized",
];

// The default Tiller layout: Description and Category happen to be adjacent.
const TILLER_HEADERS = [
  "Date",
  "Description",
  "Category",
  "Amount",
  "Account",
  "Full Description",
  "Transaction ID",
  "AI AutoCat",
];

// A layout where an ARRAYFORMULA-driven column sits between the two columns we
// write. Batching must not span it.
const HEADERS_WITH_FORMULA_GAP = [
  "Date",
  "Description",
  "Month",            // <- ARRAYFORMULA column, must never be written
  "Category",
  "Amount",
  "Full Description",
  "Transaction ID",
  "AI AutoCat",
];

// Columns deliberately jumbled, to prove nothing assumes a position.
const SHUFFLED_HEADERS = [
  "AI AutoCat",
  "Transaction ID",
  "Amount",
  "Category",
  "Full Description",
  "Date",
  "Description",
];

// Values returned from the vm context carry that context's prototypes, which
// deepStrictEqual treats as a mismatch. Re-home them in this realm first.
function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function buildRows(headers, transactionIds) {
  const idCol = headers.indexOf("Transaction ID");
  const fullDescCol = headers.indexOf("Full Description");
  return transactionIds.map((id, i) => {
    const row = new Array(headers.length).fill("");
    row[idCol] = id;
    if (fullDescCol !== -1) row[fullDescCol] = "RAW DESCRIPTION " + i;
    return row;
  });
}

function setup(source, headers, transactionIds) {
  const sheet = new FakeSheet("Transactions", headers, buildRows(headers, transactionIds));
  const categorySheet = new FakeSheet(
    "Categories",
    ["Category", "Group"],
    CATEGORIES.map((c) => [c, "G"])
  );
  const spreadsheet = new FakeSpreadsheet({
    Transactions: sheet,
    Categories: categorySheet,
  });
  const loaded = loadScripts([source], { spreadsheet });
  return { sheet, categorySheet, spreadsheet, ...loaded };
}

/** Flattens recorded writes into a comparable {a1, value} list. */
function cellWrites(sheet) {
  return sheet.writes.map((w) => ({ a1: w.a1, value: w.value }));
}

function sortWrites(writes) {
  return writes
    .slice()
    .sort((a, b) => (a.a1 < b.a1 ? -1 : a.a1 > b.a1 ? 1 : 0));
}

const currentSource = readWorkingTree("ai_autocat.gs");
const baselineSource = readAtRevision("ai_autocat.gs", BASELINE_REV);
// Upstream rewrote writeUpdatedTransactions in ec84bce with its own batching.
// This branch keeps that approach and adds the write planner on top, so the
// resulting sheet must match upstream's exactly.
const UPSTREAM_REV = process.env.UPSTREAM_REV || "ec84bce";
const upstreamSource = readAtRevision("ai_autocat.gs", UPSTREAM_REV);

const UPDATES = [
  { transaction_id: "TXN-001", updated_description: "Safeway", category: "Groceries" },
  { transaction_id: "TXN-003", updated_description: "Blue Bottle Coffee", category: "Restaurants" },
  { transaction_id: "TXN-004", updated_description: "PG&E", category: "Utilities" },
];

test("writes the same cells and values as the original implementation", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004", "TXN-005"];

  const before = setup(baselineSource, TILLER_HEADERS, ids);
  before.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  const after = setup(currentSource, TILLER_HEADERS, ids);
  after.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  assert.deepStrictEqual(
    sortWrites(cellWrites(after.sheet)),
    sortWrites(cellWrites(before.sheet))
  );
});

test("leaves the sheet in the same state as upstream's implementation", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004", "TXN-005"];

  const upstream = setup(upstreamSource, TILLER_HEADERS, ids);
  upstream.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  const after = setup(currentSource, TILLER_HEADERS, ids);
  after.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  // The resulting data is what matters, not how many calls it took.
  assert.deepStrictEqual(plain(after.sheet.data), plain(upstream.sheet.data));
});

test("touches no more cells than upstream, and never a different one", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004", "TXN-005"];

  const upstream = setup(upstreamSource, TILLER_HEADERS, ids);
  upstream.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  const after = setup(currentSource, TILLER_HEADERS, ids);
  after.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  const upstreamCells = new Set(upstream.sheet.writes.map((w) => w.a1));
  const ourCells = new Set(after.sheet.writes.map((w) => w.a1));

  for (const a1 of ourCells) {
    assert.ok(
      upstreamCells.has(a1),
      "wrote " + a1 + ", which upstream does not touch"
    );
  }
});

test("does not rewrite a description the model did not supply", () => {
  const ids = ["TXN-001", "TXN-002"];
  const { sheet, context } = setup(currentSource, TILLER_HEADERS, ids);

  const descCol = columnIndexToLetter(TILLER_HEADERS.indexOf("Description"));

  context.writeUpdatedTransactions(
    [
      { transaction_id: "TXN-001", updated_description: "", category: "Groceries" },
      { transaction_id: "TXN-002", updated_description: "Safeway", category: "Groceries" },
    ],
    CATEGORIES
  );

  // An empty description must leave the cell alone entirely. Rewriting it - even
  // with its current value - would turn an ARRAYFORMULA-derived cell into a
  // literal, which is invariant (b).
  const touchedDescCells = sheet.writes
    .filter((w) => w.column === "Description")
    .map((w) => w.a1);

  assert.deepStrictEqual(touchedDescCells, [descCol + "3"]);
});

test("stops reading the ID column once every transaction is found", () => {
  // 3000 rows, with the targets at the top - the common case after a sync.
  const ids = [];
  for (let i = 0; i < 3000; i++) ids.push("TXN-" + String(i).padStart(4, "0"));

  const { sheet, context } = setup(currentSource, TILLER_HEADERS, ids);

  const reads = [];
  const originalGetRange = sheet.getRange.bind(sheet);
  sheet.getRange = function (a1OrRow, col, numRows, numCols) {
    if (typeof a1OrRow === "number" && numRows > 1) {
      reads.push(numRows);
    }
    return originalGetRange(a1OrRow, col, numRows, numCols);
  };

  context.writeUpdatedTransactions(
    [
      { transaction_id: "TXN-0000", updated_description: "A", category: "Groceries" },
      { transaction_id: "TXN-0005", updated_description: "B", category: "Groceries" },
    ],
    CATEGORIES
  );

  const rowsRead = reads.reduce((a, b) => a + b, 0);
  assert.ok(
    rowsRead <= 500,
    "should have stopped after the first batch, read " + rowsRead + " rows"
  );
});

test("resolves columns by header regardless of their order", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004"];

  const shuffled = setup(currentSource, SHUFFLED_HEADERS, ids);
  shuffled.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  // Every write must land in one of the three columns we own, identified by
  // header name rather than index.
  const written = shuffled.sheet.writes;
  assert.ok(written.length > 0, "expected writes");
  for (const w of written) {
    assert.ok(
      ["Description", "Category", "AI AutoCat"].includes(w.column),
      "wrote into unexpected column " + w.column + " at " + w.a1
    );
  }

  // And the values must match the update they came from.
  const descCol = columnIndexToLetter(SHUFFLED_HEADERS.indexOf("Description"));
  const catCol = columnIndexToLetter(SHUFFLED_HEADERS.indexOf("Category"));
  const rowFor = { "TXN-001": 2, "TXN-003": 4, "TXN-004": 5 };

  for (const update of UPDATES) {
    const row = rowFor[update.transaction_id];
    const descWrite = written.find((w) => w.a1 === descCol + row);
    const catWrite = written.find((w) => w.a1 === catCol + row);
    assert.strictEqual(descWrite.value, update.updated_description);
    assert.strictEqual(catWrite.value, update.category);
  }
});

test("never writes into a column it does not own, even when batching", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004", "TXN-005"];
  const { sheet, context } = setup(currentSource, HEADERS_WITH_FORMULA_GAP, ids);

  context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  // "Month" sits between Description and Category. If the planner merged those
  // into one range it would clobber the ARRAYFORMULA driving this column.
  const clobbered = sheet.writes.filter((w) => w.column === "Month");
  assert.deepStrictEqual(
    clobbered,
    [],
    "wrote into the ARRAYFORMULA column: " + JSON.stringify(clobbered)
  );

  const ownedColumns = new Set(["Description", "Category", "AI AutoCat"]);
  for (const w of sheet.writes) {
    assert.ok(ownedColumns.has(w.column), "unexpected write to " + w.column);
  }
});

test("never writes a row that was not part of the update set", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004", "TXN-005"];
  const { sheet, context } = setup(currentSource, TILLER_HEADERS, ids);

  context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  // TXN-002 and TXN-005 are rows 3 and 6 and were not updated. A vertical
  // batch must not span them.
  const untouchedRows = new Set([1, 3, 6]);
  for (const w of sheet.writes) {
    assert.ok(
      !untouchedRows.has(w.row),
      "wrote to row " + w.row + " which was not in the update set (" + w.a1 + ")"
    );
  }
});

test("missing optional AI AutoCat column is skipped cleanly", () => {
  const headers = TILLER_HEADERS.filter((h) => h !== "AI AutoCat");
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004"];

  const after = setup(currentSource, headers, ids);
  after.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  // The new guard tests truthiness, so nothing is attempted and nothing is
  // logged. The old code compared against null, built the range "2", and threw
  // once per transaction into a swallowed catch.
  const errorLogs = after.logs.filter((l) => l instanceof Error || (l && l.message));
  assert.deepStrictEqual(errorLogs, [], "expected no swallowed range errors");

  for (const w of after.sheet.writes) {
    assert.ok(["Description", "Category"].includes(w.column));
  }

  // Description and Category writes still match the old implementation exactly.
  const before = setup(baselineSource, headers, ids);
  before.context.writeUpdatedTransactions(UPDATES, CATEGORIES);

  assert.deepStrictEqual(
    sortWrites(cellWrites(after.sheet)),
    sortWrites(cellWrites(before.sheet).filter((w) => w.value !== "TRUE"))
  );
});

test("falls back when the model returns a category not on the allowed list", () => {
  const ids = ["TXN-001"];
  const { sheet, context } = setup(currentSource, TILLER_HEADERS, ids);

  context.writeUpdatedTransactions(
    [{ transaction_id: "TXN-001", updated_description: "Mystery", category: "Not A Real Category" }],
    CATEGORIES
  );

  const catCol = columnIndexToLetter(TILLER_HEADERS.indexOf("Category"));
  const catWrite = sheet.writes.find((w) => w.a1 === catCol + "2");
  assert.strictEqual(catWrite.value, "To Be Categorized");
});

test("unknown transaction ids are ignored", () => {
  const ids = ["TXN-001", "TXN-002"];
  const { sheet, context } = setup(currentSource, TILLER_HEADERS, ids);

  context.writeUpdatedTransactions(
    [{ transaction_id: "TXN-DOES-NOT-EXIST", updated_description: "x", category: "Groceries" }],
    CATEGORIES
  );

  assert.deepStrictEqual(plain(sheet.writes), []);
});

test("duplicate transaction ids resolve to the first matching row", () => {
  const ids = ["TXN-DUP", "TXN-OTHER", "TXN-DUP"];
  const { sheet, context } = setup(currentSource, TILLER_HEADERS, ids);

  context.writeUpdatedTransactions(
    [{ transaction_id: "TXN-DUP", updated_description: "First", category: "Groceries" }],
    CATEGORIES
  );

  const rows = new Set(sheet.writes.map((w) => w.row));
  assert.deepStrictEqual([...rows], [2], "should write only the first matching row");
});

// --- The planner, tested directly -------------------------------------------

function coveredCells(blocks) {
  const cells = [];
  for (const b of blocks) {
    for (let r = 0; r < b.values.length; r++) {
      for (let c = 0; c < b.values[r].length; c++) {
        cells.push({
          row: b.row + r,
          column: b.column + c,
          value: b.values[r][c],
        });
      }
    }
  }
  return cells;
}

function key(cell) {
  return cell.row + ":" + cell.column;
}

test("planCellWrites covers exactly the requested cells and nothing else", () => {
  const { context } = setup(currentSource, TILLER_HEADERS, ["TXN-001"]);
  const planCellWrites = context.planCellWrites;

  // Deterministic pseudo-random sweep over many shapes.
  let seed = 12345;
  const rand = (n) => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed % n;
  };

  for (let iteration = 0; iteration < 500; iteration++) {
    const requested = [];
    const seen = new Set();
    const count = 1 + rand(12);
    for (let i = 0; i < count; i++) {
      const cell = { row: 2 + rand(8), column: 1 + rand(8), value: "v" + i };
      if (seen.has(key(cell))) continue;
      seen.add(key(cell));
      requested.push(cell);
    }

    const blocks = plain(planCellWrites(requested));
    const covered = coveredCells(blocks);

    assert.strictEqual(
      covered.length,
      requested.length,
      "plan covered " + covered.length + " cells for " + requested.length + " requests"
    );

    const coveredKeys = new Set(covered.map(key));
    for (const cell of requested) {
      assert.ok(coveredKeys.has(key(cell)), "missing cell " + key(cell));
    }
    for (const cell of covered) {
      const match = requested.find((r) => key(r) === key(cell));
      assert.ok(match, "plan invented cell " + key(cell));
      assert.strictEqual(cell.value, match.value, "wrong value at " + key(cell));
    }
  }
});

test("planCellWrites merges adjacent cells in a row into one block", () => {
  const { context } = setup(currentSource, TILLER_HEADERS, ["TXN-001"]);
  const blocks = context.planCellWrites([
    { row: 2, column: 2, value: "desc" },
    { row: 2, column: 3, value: "cat" },
  ]);

  assert.strictEqual(blocks.length, 1);
  assert.deepStrictEqual(plain(blocks[0].values), [["desc", "cat"]]);
});

test("planCellWrites does not merge across a gap", () => {
  const { context } = setup(currentSource, TILLER_HEADERS, ["TXN-001"]);
  const blocks = context.planCellWrites([
    { row: 2, column: 2, value: "desc" },
    { row: 2, column: 4, value: "cat" }, // column 3 is the formula column
  ]);

  assert.strictEqual(blocks.length, 2, "a gap must split the write");
  const columns = blocks.map((b) => b.column).sort();
  assert.deepStrictEqual(plain(columns), [2, 4]);
});

test("planCellWrites merges contiguous rows with the same column span", () => {
  const { context } = setup(currentSource, TILLER_HEADERS, ["TXN-001"]);
  const blocks = context.planCellWrites([
    { row: 2, column: 2, value: "d1" },
    { row: 2, column: 3, value: "c1" },
    { row: 3, column: 2, value: "d2" },
    { row: 3, column: 3, value: "c2" },
    { row: 4, column: 2, value: "d3" },
    { row: 4, column: 3, value: "c3" },
  ]);

  assert.strictEqual(blocks.length, 1, "a contiguous rectangle should be one write");
  assert.deepStrictEqual(plain(blocks[0]), {
    row: 2,
    column: 2,
    values: [
      ["d1", "c1"],
      ["d2", "c2"],
      ["d3", "c3"],
    ],
  });
});

test("planCellWrites does not merge rows that are not contiguous", () => {
  const { context } = setup(currentSource, TILLER_HEADERS, ["TXN-001"]);
  const blocks = context.planCellWrites([
    { row: 2, column: 2, value: "d1" },
    { row: 2, column: 3, value: "c1" },
    { row: 5, column: 2, value: "d2" }, // rows 3 and 4 must not be touched
    { row: 5, column: 3, value: "c2" },
  ]);

  assert.strictEqual(blocks.length, 2);
});

// --- ARRAYFORMULA preservation ----------------------------------------------
//
// The hazard invariant (b) exists for. Upstream's implementation reads each
// contiguous range with getValues() and writes the whole range back with
// setValues(). For a transaction whose updated_description is empty, that reads
// the existing value and writes it straight back - and if the cell was derived
// from a sheet-level ARRAYFORMULA, getValues() returns the computed value while
// setValues() stores it as a literal, silently destroying the formula.

function setupWithFormulaDescription(source, ids) {
  const ctx = setup(source, TILLER_HEADERS, ids);
  const descCol = columnIndexToLetter(TILLER_HEADERS.indexOf("Description"));

  // Description is driven by an ARRAYFORMULA across the data rows.
  ctx.sheet.markFormulaDerived(ids.map((_, i) => descCol + (i + 2)));
  return ctx;
}

test("does not break an ARRAYFORMULA when the model returns no description", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003"];
  const updates = [
    { transaction_id: "TXN-001", updated_description: "", category: "Groceries" },
    { transaction_id: "TXN-002", updated_description: "", category: "Utilities" },
  ];

  const ours = setupWithFormulaDescription(currentSource, ids);
  ours.context.writeUpdatedTransactions(updates, CATEGORIES);

  assert.deepStrictEqual(
    plain(ours.sheet.brokenFormulas),
    [],
    "wrote into ARRAYFORMULA-derived description cells"
  );

  // The categories still land, so skipping the description costs nothing.
  const catCol = columnIndexToLetter(TILLER_HEADERS.indexOf("Category"));
  const catWrites = ours.sheet.writes
    .filter((w) => w.column === "Category")
    .map((w) => w.a1)
    .sort();
  assert.deepStrictEqual(catWrites, [catCol + "2", catCol + "3"]);
});

test("upstream's read-modify-write breaks that ARRAYFORMULA", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003"];
  const updates = [
    { transaction_id: "TXN-001", updated_description: "", category: "Groceries" },
    { transaction_id: "TXN-002", updated_description: "", category: "Utilities" },
  ];

  const upstream = setupWithFormulaDescription(upstreamSource, ids);
  upstream.context.writeUpdatedTransactions(updates, CATEGORIES);

  // Documents the regression this branch avoids. If upstream later stops
  // rewriting untouched description cells, this test should be deleted.
  assert.ok(
    upstream.sheet.brokenFormulas.length > 0,
    "expected upstream to rewrite the derived description cells"
  );
  assert.deepStrictEqual(
    plain(upstream.sheet.brokenFormulas.map((b) => b.a1)).sort(),
    ["B2", "B3"]
  );
});

test("a formula column between the written columns survives both row batching and column batching", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004"];
  const ctx = setup(currentSource, HEADERS_WITH_FORMULA_GAP, ids);

  const monthCol = columnIndexToLetter(HEADERS_WITH_FORMULA_GAP.indexOf("Month"));
  ctx.sheet.markFormulaDerived(ids.map((_, i) => monthCol + (i + 2)));

  // All four rows updated and contiguous: the planner will merge vertically,
  // which is exactly when an over-wide block would span the Month column.
  ctx.context.writeUpdatedTransactions(
    ids.map((id, i) => ({
      transaction_id: id,
      updated_description: "Merchant " + i,
      category: "Groceries",
    })),
    CATEGORIES
  );

  assert.deepStrictEqual(plain(ctx.sheet.brokenFormulas), []);
});

test("rows between non-contiguous updates keep their formulas", () => {
  const ids = ["TXN-001", "TXN-002", "TXN-003", "TXN-004", "TXN-005"];
  const ctx = setup(currentSource, TILLER_HEADERS, ids);

  // Rows 3 and 5 are not updated; mark their Description/Category as derived.
  const descCol = columnIndexToLetter(TILLER_HEADERS.indexOf("Description"));
  const catCol = columnIndexToLetter(TILLER_HEADERS.indexOf("Category"));
  ctx.sheet.markFormulaDerived([
    descCol + "3", catCol + "3",
    descCol + "5", catCol + "5",
  ]);

  ctx.context.writeUpdatedTransactions(
    [
      { transaction_id: "TXN-001", updated_description: "A", category: "Groceries" },
      { transaction_id: "TXN-003", updated_description: "B", category: "Groceries" },
      { transaction_id: "TXN-005", updated_description: "C", category: "Groceries" },
    ],
    CATEGORIES
  );

  assert.deepStrictEqual(
    plain(ctx.sheet.brokenFormulas),
    [],
    "a vertical merge spanned a row that was not being updated"
  );
});
