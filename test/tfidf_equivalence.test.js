/**
 * Characterisation tests for the TF-IDF scoring refactor.
 *
 * These load the committed version of tfidf_search.gs and the working tree
 * version side by side and assert they produce identical rankings and identical
 * scores. The refactor is a pure performance change (precomputed term
 * frequencies instead of re-tokenising per comparison), so any divergence here
 * is a regression.
 */
const test = require("node:test");
const assert = require("node:assert");
const { loadScripts, readWorkingTree, readAtRevision } = require("./harness");

// The implementation these tests characterise, i.e. the last commit before the
// Vertex AI / performance work. Pinned to a SHA rather than HEAD so the
// comparison stays meaningful once these changes are themselves committed.
const BASELINE_REV = process.env.BASELINE_REV || "9ac70da1025494e53d9bdd16595c400cf87dd4bf";

// Values coming back from a vm context carry that context's prototypes, so
// deepStrictEqual would reject them even when the contents match. Round-tripping
// through JSON re-homes them in this realm.
function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function makeSearcher(source, documents, options) {
  const { context } = loadScripts([source]);
  return new context.TFIDFSearch(documents, options);
}

// A corpus shaped like real Tiller data: repeated merchants, shared tokens,
// varying lengths, plus deliberately awkward rows.
const DOCUMENTS = [
  { id: "t1", text: "AMAZON MKTPLACE PMTS AMZN COM BILL WA", updatedText: "Amazon", category: "Shopping", date: new Date("2024-01-05"), amount: -42.1 },
  { id: "t2", text: "AMAZON MKTPLACE PMTS AMZN COM BILL WA", updatedText: "Amazon", category: "Shopping", date: new Date("2024-03-11"), amount: -18.0 },
  { id: "t3", text: "AMZN Mktp US*2K4LM9XY3", updatedText: "Amazon", category: "Shopping", date: new Date("2024-02-02"), amount: -7.25 },
  { id: "t4", text: "TST* STONEMILL MATCHA SAN FRANCISCOCA", updatedText: "Stonemill Matcha", category: "Restaurants", date: new Date("2024-02-14"), amount: -23.5 },
  { id: "t5", text: "SQ *BLUE BOTTLE COFFEE OAKLAND CA", updatedText: "Blue Bottle Coffee", category: "Coffee Shops", date: new Date("2024-02-15"), amount: -6.75 },
  { id: "t6", text: "PG&E WEB ONLINE PAYMENT 9284712", updatedText: "PG&E", category: "Utilities", date: new Date("2024-01-20"), amount: -180.44 },
  { id: "t7", text: "PG&E WEB ONLINE PAYMENT 1093822", updatedText: "PG&E", category: "Utilities", date: new Date("2024-02-20"), amount: -175.02 },
  { id: "t8", text: "SAFEWAY #1234 SAN FRANCISCO CA", updatedText: "Safeway", category: "Groceries", date: new Date("2024-03-01"), amount: -88.13 },
  { id: "t9", text: "SAFEWAY #5678 OAKLAND CA", updatedText: "Safeway", category: "Groceries", date: new Date("2024-03-02"), amount: -55.9 },
  { id: "t10", text: "UNITED AIRLINES 0162381729301 SAN FRANCISCO", updatedText: "United Airlines", category: "Travel", date: new Date("2024-04-01"), amount: -612.3 },
  // Awkward rows: empty text, punctuation only, and tokens shorter than minTermSize.
  { id: "t11", text: "", updatedText: "", category: "To Be Categorized", date: new Date("2024-01-01"), amount: 0 },
  { id: "t12", text: "-- ** ..", updatedText: "", category: "To Be Categorized", date: new Date("2024-01-02"), amount: 0 },
  { id: "t13", text: "a an to of it", updatedText: "", category: "To Be Categorized", date: new Date("2024-01-03"), amount: 0 },
  // Identical text and identical dates, to exercise the tie-break path.
  { id: "t14", text: "NETFLIX COM LOS GATOS CA", updatedText: "Netflix", category: "Entertainment", date: new Date("2024-05-01"), amount: -15.49 },
  { id: "t15", text: "NETFLIX COM LOS GATOS CA", updatedText: "Netflix", category: "Entertainment", date: new Date("2024-05-01"), amount: -15.49 },
];

const QUERIES = [
  // Ordinary lookups.
  "AMAZON MKTPLACE PMTS AMZN COM BILL WA",
  "TST* STONEMILL MATCHA SAN FRANCISCOCA",
  "SAFEWAY #9999 BERKELEY CA",
  "PG&E WEB ONLINE PAYMENT 5555555",
  "NETFLIX COM LOS GATOS CA",
  // Repeated terms: these must boost the score twice but count once toward the
  // distinct-term multiplier. This is the exact case the old wordScores map
  // handled implicitly by keying on the word.
  "AMAZON AMAZON AMAZON MKTPLACE",
  "SAFEWAY SAFEWAY",
  "NETFLIX NETFLIX NETFLIX NETFLIX COM COM",
  // Terms that match nothing at all.
  "ZZZZQQQQ NONEXISTENT MERCHANT",
  "QQQQ",
  // Mixed: some terms match, some do not.
  "AMAZON ZZZZQQQQ SAFEWAY NONEXISTENT",
  // Degenerate queries.
  "",
  "   ",
  "a an to of",
  "-- ** ..",
  "ab cd ef",
  // Case and punctuation variation.
  "amazon mktplace pmts",
  "Amazon's Mktplace",
  "safeway---oakland",
];

const OPTION_SETS = [
  { useStopWords: true, matchThreshold: 0.25, minTermSize: 3 },
  { useStopWords: true, matchThreshold: 0.0, minTermSize: 3 },
  { useStopWords: false, matchThreshold: 0.0, minTermSize: 2 },
  { useStopWords: true, matchThreshold: 0.5, minTermSize: 4 },
];

// Fields that make up the observable contract of a search result. `details` is
// deliberately excluded: it was debug-only scaffolding and the refactor removes
// it from the hot path.
function projectResult(r) {
  return {
    id: r.id,
    text: r.text,
    updatedText: r.updatedText,
    category: r.category,
    amount: r.amount,
    // Scores are compared exactly. The refactor reorders arithmetic only in
    // ways that are bit-identical, so this should hold without a tolerance.
    score: Number.isNaN(r.score) ? "NaN" : r.score,
    date: r.date ? r.date.getTime() : null,
  };
}

const baselineSource = readAtRevision("tfidf_search.gs", BASELINE_REV);
const currentSource = readWorkingTree("tfidf_search.gs");

test("search() returns identical results before and after the refactor", () => {
  let comparisons = 0;

  for (const options of OPTION_SETS) {
    const before = makeSearcher(baselineSource, DOCUMENTS, options);
    const after = makeSearcher(currentSource, DOCUMENTS, options);

    for (const query of QUERIES) {
      for (const limit of [1, 3, 5, 50]) {
        const expected = plain(before.search(query, limit).map(projectResult));
        const actual = plain(after.search(query, limit).map(projectResult));

        assert.deepStrictEqual(
          actual,
          expected,
          "divergence for query " +
            JSON.stringify(query) +
            " limit=" + limit +
            " options=" + JSON.stringify(options)
        );
        comparisons++;
      }
    }
  }

  assert.ok(comparisons > 200, "expected a broad comparison sweep");
});

test("repeated query terms score the same but boost the distinct-term count once", () => {
  const options = { useStopWords: true, matchThreshold: 0.0, minTermSize: 3 };
  const before = makeSearcher(baselineSource, DOCUMENTS, options);
  const after = makeSearcher(currentSource, DOCUMENTS, options);

  // "AMAZON AMAZON" must not be treated as two distinct matching terms.
  const single = after.search("AMAZON", 5);
  const doubled = after.search("AMAZON AMAZON", 5);

  assert.deepStrictEqual(
    plain(doubled.map(projectResult)),
    plain(before.search("AMAZON AMAZON", 5).map(projectResult))
  );

  // Same ranking, and the doubled query scores exactly twice the single one:
  // the term contributes twice, while the 1.2^(distinct-1) multiplier is
  // unchanged at 1.2^0.
  assert.deepStrictEqual(
    plain(doubled.map((r) => r.id)),
    plain(single.map((r) => r.id))
  );
  for (let i = 0; i < single.length; i++) {
    assert.ok(
      Math.abs(doubled[i].score - single[i].score * 2) < 1e-12,
      "repeated term should double the score, not change the term-count boost"
    );
  }
});

test("documents that tokenise to nothing behave identically", () => {
  const options = { useStopWords: true, matchThreshold: 0.0, minTermSize: 3 };
  const before = makeSearcher(baselineSource, DOCUMENTS, options);
  const after = makeSearcher(currentSource, DOCUMENTS, options);

  // t11/t12/t13 tokenise to zero terms, so tf() divides by zero. Whatever the
  // original did with the resulting NaN, the refactor must do the same.
  const emptyDocIds = ["t11", "t12", "t13"];
  const expected = plain(before.search("AMAZON", 50).map((r) => r.id));
  const actual = plain(after.search("AMAZON", 50).map((r) => r.id));

  assert.deepStrictEqual(actual, expected);
  for (const id of emptyDocIds) {
    assert.ok(
      !actual.includes(id),
      "zero-term document " + id + " should not surface as a match"
    );
  }
});

test("wordDocs index is built identically", () => {
  const options = { useStopWords: true, matchThreshold: 0.25, minTermSize: 3 };
  const before = makeSearcher(baselineSource, DOCUMENTS, options);
  const after = makeSearcher(currentSource, DOCUMENTS, options);

  assert.deepStrictEqual(
    plain(Object.keys(after.wordDocs).sort()),
    plain(Object.keys(before.wordDocs).sort())
  );
  for (const word of Object.keys(before.wordDocs)) {
    assert.deepStrictEqual(
      plain(after.wordDocs[word]),
      plain(before.wordDocs[word]),
      "document list for term " + JSON.stringify(word) + " differs"
    );
  }
});

test("tokenize() is unchanged", () => {
  const options = { useStopWords: true, matchThreshold: 0.25, minTermSize: 3 };
  const before = makeSearcher(baselineSource, [], options);
  const after = makeSearcher(currentSource, [], options);

  const samples = QUERIES.concat(DOCUMENTS.map((d) => d.text));
  for (const s of samples) {
    assert.deepStrictEqual(
      plain(after.tokenize(s)),
      plain(before.tokenize(s)),
      "tokenisation differs for " + JSON.stringify(s)
    );
  }
});
