/**
 * Tests for the search index options, the category list, and the performance
 * characteristic the refactor was for.
 */
const test = require("node:test");
const assert = require("node:assert");
const {
  FakeSheet,
  FakeSpreadsheet,
  loadScripts,
  readWorkingTree,
  readAtRevision,
} = require("./harness");

// The implementation these tests characterise, i.e. the last commit before the
// Vertex AI / performance work. Pinned to a SHA rather than HEAD so the
// comparison stays meaningful once these changes are themselves committed.
const BASELINE_REV = process.env.BASELINE_REV || "9ac70da1025494e53d9bdd16595c400cf87dd4bf";

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

const currentTfidf = readWorkingTree("tfidf_search.gs");
const baselineTfidf = readAtRevision("tfidf_search.gs", BASELINE_REV);
const currentAutocat = readWorkingTree("ai_autocat.gs");
const baselineAutocat = readAtRevision("ai_autocat.gs", BASELINE_REV);

// --- Category list ----------------------------------------------------------

function categoriesSetup(autocatSource) {
  // A Categories sheet with only a handful of real rows. The sheet itself is
  // 1000 rows tall, which is why the open-ended read returns mostly blanks.
  const categorySheet = new FakeSheet(
    "Categories",
    ["Category", "Group", "Type"],
    [
      ["Groceries", "Food", "Expense"],
      ["Restaurants", "Food", "Expense"],
      ["Utilities", "Home", "Expense"],
      ["To Be Categorized", "Other", "Expense"],
    ],
    1000
  );
  const spreadsheet = new FakeSpreadsheet({
    Categories: categorySheet,
    Transactions: new FakeSheet("Transactions", ["Transaction ID"], []),
  });
  return loadScripts([autocatSource], { spreadsheet });
}

test("getAllowedCategories returns only the real categories", () => {
  const { context } = categoriesSetup(currentAutocat);
  const categories = plain(context.getAllowedCategories());

  assert.deepStrictEqual(categories, [
    "Groceries",
    "Restaurants",
    "Utilities",
    "To Be Categorized",
  ]);
});

test("blank category rows are no longer padded into the model prompt", () => {
  const before = categoriesSetup(baselineAutocat);
  const after = categoriesSetup(currentAutocat);

  const beforeList = plain(before.context.getAllowedCategories());
  const afterList = plain(after.context.getAllowedCategories());

  // The old behaviour dragged the full height of the sheet into the prompt.
  assert.ok(
    beforeList.length > 900,
    "expected the old implementation to include blank rows, got " + beforeList.length
  );
  assert.strictEqual(beforeList.filter((c) => c === "").length, beforeList.length - 4);

  // The new one carries only real values, and preserves them in sheet order.
  assert.strictEqual(afterList.length, 4);
  assert.deepStrictEqual(
    afterList,
    beforeList.filter((c) => c !== "")
  );
});

// --- Search index options ---------------------------------------------------

function indexSetup(tfidfSource, autocatSource) {
  const transactions = new FakeSheet(
    "Transactions",
    [
      "Date",
      "Description",
      "Category",
      "Amount",
      "Full Description",
      "Transaction ID",
    ],
    [
      ["2024-01-01", "Safeway", "Groceries", -20, "SAFEWAY #1234 SF CA", "T1"],
      ["2024-01-02", "Amazon", "Shopping", -35, "AMAZON MKTPLACE PMTS", "T2"],
    ]
  );
  const spreadsheet = new FakeSpreadsheet({ Transactions: transactions });
  return loadScripts([tfidfSource, autocatSource], { spreadsheet });
}

test("createSearchIndex honours the options it is handed", () => {
  const { context } = indexSetup(currentTfidf, currentAutocat);

  const searcher = context.createSearchIndex(
    "Transactions",
    "F", "E", "B", "A", "C", "D",
    2,
    { minTermSize: 5, matchThreshold: 0.9, useStopWords: false }
  );

  assert.strictEqual(searcher.minTermSize, 5);
  assert.strictEqual(searcher.matchThreshold, 0.9);
  assert.strictEqual(searcher.useStopWords, false);
});

test("options passed by createSearchIndexWithStandardColumns were previously dropped", () => {
  const before = indexSetup(baselineTfidf, baselineAutocat);
  const after = indexSetup(currentTfidf, currentAutocat);

  const beforeSearcher = before.context.createSearchIndex(
    "Transactions", "F", "E", "B", "A", "C", "D", 2, { minTermSize: 7 }
  );
  const afterSearcher = after.context.createSearchIndex(
    "Transactions", "F", "E", "B", "A", "C", "D", 2, { minTermSize: 7 }
  );

  // Documents the bug this change fixes: the old signature had no options
  // parameter, so the argument was silently discarded.
  assert.strictEqual(beforeSearcher.minTermSize, 3, "baseline ignored the option");
  assert.strictEqual(afterSearcher.minTermSize, 7, "options are now honoured");
});

test("the default options used in production are unchanged", () => {
  const before = indexSetup(baselineTfidf, baselineAutocat);
  const after = indexSetup(currentTfidf, currentAutocat);

  // createSearchIndexWithStandardColumns passes { minTermSize: 3 }, which
  // happens to equal the old hardcoded value - so real behaviour is identical.
  const beforeSearcher = before.context.createSearchIndexWithStandardColumns({ minTermSize: 3 });
  const afterSearcher = after.context.createSearchIndexWithStandardColumns({ minTermSize: 3 });

  assert.strictEqual(afterSearcher.minTermSize, beforeSearcher.minTermSize);
  assert.strictEqual(afterSearcher.matchThreshold, beforeSearcher.matchThreshold);
  assert.strictEqual(afterSearcher.useStopWords, beforeSearcher.useStopWords);
});

// --- Performance ------------------------------------------------------------

function buildCorpus(size) {
  const merchants = [
    "SAFEWAY #1234 SAN FRANCISCO CA",
    "AMAZON MKTPLACE PMTS AMZN COM BILL WA",
    "SQ *BLUE BOTTLE COFFEE OAKLAND CA",
    "PG&E WEB ONLINE PAYMENT 9284712",
    "NETFLIX COM LOS GATOS CA",
  ];
  const docs = [];
  for (let i = 0; i < size; i++) {
    docs.push({
      id: "T" + i,
      text: merchants[i % merchants.length] + " " + i,
      updatedText: "M" + (i % merchants.length),
      category: "Groceries",
      date: new Date(2024, 0, 1 + (i % 300)),
      amount: -(i % 100),
    });
  }
  return docs;
}

/** Counts tokenize() calls during a search, which is the cost the refactor targets. */
function countTokenizeCalls(source, documents, query) {
  const { context } = loadScripts([source]);
  const searcher = new context.TFIDFSearch(documents, {
    useStopWords: true,
    matchThreshold: 0.25,
    minTermSize: 3,
  });

  let calls = 0;
  const original = searcher.tokenize.bind(searcher);
  searcher.tokenize = function (text) {
    calls++;
    return original(text);
  };

  searcher.search(query, 3);
  return calls;
}

test("search no longer re-tokenises every document for every query term", () => {
  const documents = buildCorpus(400);
  const query = "SAFEWAY SAN FRANCISCO CA 1234";

  const beforeCalls = countTokenizeCalls(baselineTfidf, documents, query);
  const afterCalls = countTokenizeCalls(currentTfidf, documents, query);

  // The old implementation tokenised each document once per query term, so the
  // count scaled with the corpus. The new one tokenises only the query.
  assert.ok(
    beforeCalls > 1000,
    "expected the baseline to scale with the corpus, saw " + beforeCalls
  );
  assert.strictEqual(
    afterCalls,
    1,
    "search should tokenise the query and nothing else, saw " + afterCalls
  );
});

test("index build cost does not regress", () => {
  const documents = buildCorpus(2000);

  const before = loadScripts([baselineTfidf]);
  const after = loadScripts([currentTfidf]);

  const opts = { useStopWords: true, matchThreshold: 0.25, minTermSize: 3 };
  const beforeSearcher = new before.context.TFIDFSearch(documents, opts);
  const afterSearcher = new after.context.TFIDFSearch(documents, opts);

  // Both tokenise each document exactly once at build time; the new version
  // additionally records term counts, which must stay cheap.
  assert.ok(
    afterSearcher.timing.processDocuments <=
      beforeSearcher.timing.processDocuments + 250,
    "index build regressed: " +
      afterSearcher.timing.processDocuments +
      "ms vs " +
      beforeSearcher.timing.processDocuments +
      "ms"
  );
});

// --- Robustness -------------------------------------------------------------

test("terms that collide with Object prototype keys are not dropped", () => {
  // Tokens are lowercased, so the only real English words that collide with
  // Object.prototype members are "constructor" and "__proto__". Looking them up
  // on a plain object literal returns an inherited member, which the stop-word
  // filter read as truthy and discarded.
  const documents = [
    { id: "c1", text: "constructor supply company payment", updatedText: "Constructor Supply", category: "Other", date: new Date("2024-01-01"), amount: -1 },
    { id: "c2", text: "ordinary merchant payment", updatedText: "Ordinary", category: "Other", date: new Date("2024-01-02"), amount: -2 },
  ];
  const opts = { useStopWords: true, matchThreshold: 0.0, minTermSize: 3 };

  const before = loadScripts([baselineTfidf]);
  const after = loadScripts([currentTfidf]);

  const beforeSearcher = new before.context.TFIDFSearch(documents, opts);
  const afterSearcher = new after.context.TFIDFSearch(documents, opts);

  // The baseline silently swallowed the term.
  assert.ok(
    !plain(beforeSearcher.tokenize("constructor supply")).includes("constructor"),
    "expected the baseline to drop 'constructor' as a stop word"
  );
  assert.deepStrictEqual(
    plain(afterSearcher.tokenize("constructor supply")),
    ["constructor", "supply"]
  );

  // And it is now a real, searchable term.
  assert.ok(afterSearcher.wordDocs["constructor"], "term should be indexed");
  const results = afterSearcher.search("constructor supply", 5);
  assert.ok(results.length > 0, "expected matches");
  assert.strictEqual(results[0].id, "c1");

  // "__proto__" is the other collision, and must not corrupt the index.
  // Two documents, so that idf() is non-zero for a term appearing in only one.
  const protoDocs = [
    { id: "p1", text: "__proto__ merchant payment", updatedText: "P", category: "Other", date: new Date("2024-01-01"), amount: -1 },
    { id: "p2", text: "unrelated merchant payment", updatedText: "U", category: "Other", date: new Date("2024-01-02"), amount: -2 },
  ];
  const protoSearcher = new after.context.TFIDFSearch(protoDocs, opts);
  assert.ok(Array.isArray(protoSearcher.wordDocs["__proto__"]));
  assert.strictEqual(protoSearcher.search("__proto__", 5)[0].id, "p1");
});
