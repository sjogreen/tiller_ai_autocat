/**
 * Tests for the categorization pipeline: what is read from the sheet, what is
 * sent to Jev and to Gemini, how their answers combine, and what gets written.
 * Every network call is stubbed; the stubs record the requests so the tests can
 * assert exactly what data leaves the sheet.
 */
const test = require("node:test");
const assert = require("node:assert");
const {
  FakeSheet,
  FakeSpreadsheet,
  loadScripts,
  readWorkingTree,
} = require("./harness");

const SOURCES = ["gviz.gs", "tfidf_search.gs", "ai_autocat.gs"].map(readWorkingTree);

// Batch sizes as the script defines them, so the batching tests follow the
// constants rather than repeating them.
const SETTINGS = (() => {
  const { context } = loadScripts(SOURCES);
  const read = (name) => require("vm").runInContext(name, context);
  return {
    batch: read("MAX_BATCH_SIZE"),
    perRequest: read("GEMINI_ROWS_PER_REQUEST"),
    inFlight: read("GEMINI_CONCURRENCY"),
  };
})();
const B = SETTINGS.batch;
const requestsFor = (rows) => Math.ceil(rows / SETTINGS.perRequest);

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function daysAgo(n) {
  const d = new Date();
  d.setHours(12, 0, 0, 0);
  d.setDate(d.getDate() - n);
  return d;
}

function iso(d) {
  const pad = (n) => (n < 10 ? "0" : "") + n;
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
}

const HEADERS = [
  "Date",
  "Description",
  "Category",
  "Amount",
  "Full Description",
  "Transaction ID",
  "AI AutoCat",
];

const CATEGORIES = [
  ["Groceries", "Food"],
  ["Restaurants", "Food"],
  ["Utilities", "Home"],
  ["To Be Categorized", "Other"],
];

// Past rows, all categorized. One is older than the 365-day lookback.
const HISTORY = [
  [daysAgo(30), "Safeway", "Groceries", -42.1, "SAFEWAY #1234 SAN FRANCISCO CA", "H1", ""],
  [daysAgo(60), "PG&E", "Utilities", -120, "PG&E WEB ONLINE PAYMENT 9284712", "H2", ""],
  [daysAgo(90), "Mystery", "To Be Categorized", -9, "ZXQ MYSTERY CHARGE", "H3", ""],
  [daysAgo(400), "Old Diner", "Restaurants", -15, "OLD DINER OAKLAND CA", "H4", ""],
];

// Uncategorized rows, as gviz would return them: [id, full desc, amount, date].
const NEW_ROWS = [
  ["N1", "SAFEWAY #5678 SAN FRANCISCO CA", -18.5, "Date(2026,9,1)"],
  ["N2", "PG&E WEB ONLINE PAYMENT 1111111", -130, "Date(2026,9,2)"],
  ["N3", "BLUE BOTTLE COFFEE OAKLAND CA", -6, "Date(2026,9,3)"],
];

function gvizResponse(rows) {
  const table = {
    cols: rows[0].map((_, i) => ({ label: "c" + i })),
    rows: rows.map((r) => ({ c: r.map((v) => ({ v })) })),
  };
  return "/*O_o*/\ngoogle.visualization.Query.setResponse(" + JSON.stringify({ table }) + ");";
}

function response(code, body) {
  return {
    getResponseCode: () => code,
    getContentText: () => (typeof body === "string" ? body : JSON.stringify(body)),
  };
}

function geminiResponse(suggestions) {
  return response(200, {
    candidates: [
      {
        content: {
          parts: [{ text: JSON.stringify({ suggested_transactions: suggestions }) }],
        },
      },
    ],
    usageMetadata: { promptTokenCount: 100, candidatesTokenCount: 20 },
  });
}

function jevAnswer(choice, noneProbability) {
  return response(200, {
    answers: {
      precedent: {
        choice,
        probabilities: { none: noneProbability },
        confidence: 0.9,
      },
    },
    usage: { cost: 0.0001 },
  });
}

/**
 * Runs categorizeUncategorizedTransactions against a fake sheet.
 * jev(request) answers each Jev request; gemini(request) answers the Vertex call.
 */
function run({ openRouterKey, jev, gemini, newRows, headers }) {
  const hdrs = headers || HEADERS;
  const rows = HISTORY.map((r) => hdrs.map((h) => r[HEADERS.indexOf(h)] ?? "")).concat(
    (newRows || NEW_ROWS).map((n) => {
      const row = new Array(hdrs.length).fill("");
      row[hdrs.indexOf("Transaction ID")] = n[0];
      row[hdrs.indexOf("Full Description")] = n[1];
      return row;
    })
  );
  const sheet = new FakeSheet("Transactions", hdrs, rows);
  const categorySheet = new FakeSheet("Categories", ["Category", "Group"], CATEGORIES);
  const spreadsheet = new FakeSpreadsheet({
    Transactions: sheet,
    Categories: categorySheet,
  });

  const calls = { gviz: [], jev: [], gemini: [] };
  const { context, logs } = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: Object.assign(
      { GCP_PROJECT_ID: "proj" },
      openRouterKey ? { OPENROUTER_API_KEY: openRouterKey } : {}
    ),
    fetch: (url, params) => {
      if (url.includes("/gviz/")) {
        calls.gviz.push(decodeURIComponent(url));
        return {
          getContentText: () => gvizResponse(newRows || NEW_ROWS),
        };
      }
      throw new Error("Unexpected fetch " + url);
    },
    fetchAll: (requests) => {
      calls.fetchAllSizes = (calls.fetchAllSizes || []).concat([requests.length]);
      return requests.map((r) => {
        if (r.url.includes("aiplatform.googleapis.com")) {
          const request = JSON.parse(r.payload);
          calls.gemini.push({ url: r.url, request, params: r });
          return gemini(request);
        }
        assert.strictEqual(r.url, "https://openrouter.ai/api/alpha/decisions");
        const request = JSON.parse(r.payload);
        calls.jev.push({ request, headers: r.headers });
        return jev(request);
      });
    },
  });

  context.categorizeUncategorizedTransactions();

  const written = {};
  sheet.writes.forEach((w) => {
    const id = sheet.data[w.row - 1][hdrs.indexOf("Transaction ID")];
    written[id] = written[id] || {};
    written[id][w.column] = w.value;
  });
  return { calls, written, logs, context };
}

// --- Reading the sheet ------------------------------------------------------

test("isoDate handles sheet dates, gviz dates and blanks", () => {
  const { context } = loadScripts(SOURCES);
  assert.strictEqual(context.isoDate(new Date(2026, 7, 14, 15, 30)), "2026-08-14");
  assert.strictEqual(context.isoDate("Date(2026,7,14)"), "2026-08-14");
  assert.strictEqual(context.isoDate("Date(2026,0,5,0,0,0)"), "2026-01-05");
  assert.strictEqual(context.isoDate(""), "");
  assert.strictEqual(context.isoDate(null), "");
  assert.strictEqual(context.isoDate("2026-08-14"), "2026-08-14");
});

test("the uncategorized query selects amount and date, and they reach the model", () => {
  const { calls } = run({
    gemini: () => geminiResponse([]),
  });
  // Date is column A, Amount column D, Full Description E, Transaction ID F.
  assert.match(calls.gviz[0], /SELECT F, E, D, A WHERE E is not null AND C is null/);

  const sent = calls.gemini[0].request;
  const payload = JSON.parse(sent.contents[0].parts[0].text);
  assert.deepStrictEqual(
    payload.transactions.map((t) => [t.transaction_id, t.amount, t.date]),
    [
      ["N1", -18.5, "2026-10-01"],
      ["N2", -130, "2026-10-02"],
      ["N3", -6, "2026-10-03"],
    ]
  );
});

test("a sheet without Amount or Date columns still works", () => {
  const headers = ["Description", "Category", "Full Description", "Transaction ID"];
  const newRows = [["N1", "SAFEWAY #5678 SAN FRANCISCO CA"]];
  const { calls } = run({
    headers,
    newRows,
    gemini: () => geminiResponse([]),
  });
  assert.match(calls.gviz[0], /SELECT D, C WHERE/);
  const payload = JSON.parse(calls.gemini[0].request.contents[0].parts[0].text);
  assert.strictEqual("amount" in payload.transactions[0], false);
  assert.strictEqual("date" in payload.transactions[0], false);
});

test("previous transactions older than the 365-day lookback are not indexed", () => {
  const { context } = run({ gemini: () => geminiResponse([]) });
  const searcher = context.createSearchIndexWithStandardColumns({ minTermSize: 3 });
  const ids = plain(searcher.documents.map((d) => d.id));
  assert.deepStrictEqual(ids.sort(), ["H1", "H2", "H3"]);
});

// --- Without an OpenRouter key: Gemini only ---------------------------------

test("without an OpenRouter key Jev is never called and Gemini gets every row", () => {
  const { calls } = run({ gemini: () => geminiResponse([]) });
  assert.strictEqual(calls.jev.length, 0);
  const payload = JSON.parse(calls.gemini[0].request.contents[0].parts[0].text);
  assert.deepStrictEqual(
    payload.transactions.map((t) => t.transaction_id),
    ["N1", "N2", "N3"]
  );
});

test("Gemini is called with deterministic settings and a strict schema", () => {
  const { calls } = run({ gemini: () => geminiResponse([]) });
  const { request, url } = calls.gemini[0];
  assert.match(url, /publishers\/google\/models\/gemini-3\.8-flash:generateContent$/);
  assert.deepStrictEqual(
    {
      temperature: request.generationConfig.temperature,
      maxOutputTokens: request.generationConfig.maxOutputTokens,
      responseMimeType: request.generationConfig.responseMimeType,
      thinkingConfig: request.generationConfig.thinkingConfig,
    },
    {
      temperature: 0,
      maxOutputTokens: 32000,
      responseMimeType: "application/json",
      thinkingConfig: { thinkingLevel: "LOW" },
    }
  );
  assert.deepStrictEqual(
    request.generationConfig.responseSchema.properties.suggested_transactions.items.required,
    ["transaction_id", "updated_description", "category"]
  );
});

test("Gemini gets category names with groups, and previous transactions by name", () => {
  const { calls } = run({ gemini: () => geminiResponse([]) });
  const { request } = calls.gemini[0];

  // The category list is data in the user message, not baked into the prompt.
  assert.doesNotMatch(request.systemInstruction.parts[0].text, /Utilities/);

  const payload = JSON.parse(request.contents[0].parts[0].text);
  assert.deepStrictEqual(payload.allowed_categories, [
    { name: "Groceries", group: "Food" },
    { name: "Restaurants", group: "Food" },
    { name: "Utilities", group: "Home" },
    { name: "To Be Categorized", group: "Other" },
  ]);

  const n1 = payload.transactions.find((t) => t.transaction_id === "N1");
  assert.deepStrictEqual(n1, {
    transaction_id: "N1",
    original_description: "SAFEWAY #5678 SAN FRANCISCO CA",
    amount: -18.5,
    date: "2026-10-01",
    previous_transactions: [
      {
        bank_description: "SAFEWAY #1234 SAN FRANCISCO CA",
        description: "Safeway",
        category: "Groceries",
        amount: -42.1,
        date: iso(HISTORY[0][0]),
      },
    ],
  });
});

test("a declined or unknown category is written as the fallback", () => {
  const { written } = run({
    gemini: () =>
      geminiResponse([
        { transaction_id: "N1", updated_description: "Safeway", category: "Groceries" },
        { transaction_id: "N2", updated_description: "PG&E", category: "to-be-categorized" },
        { transaction_id: "N3", updated_description: "Blue Bottle Coffee", category: "Coffee" },
      ]),
  });
  assert.strictEqual(written.N1.Category, "Groceries");
  assert.strictEqual(written.N2.Category, "To Be Categorized");
  assert.strictEqual(written.N2.Description, "PG&E");
  assert.strictEqual(written.N3.Category, "To Be Categorized");
});

// --- With an OpenRouter key: Jev first --------------------------------------

test("Jev is asked only about rows with previous transactions, with raw text and names", () => {
  const { calls } = run({
    openRouterKey: "or-key",
    jev: () => jevAnswer("none", 0.9),
    gemini: () => geminiResponse([]),
  });

  // N3 (Blue Bottle) has no previous transaction, so it is not asked.
  assert.strictEqual(calls.jev.length, 2);
  assert.strictEqual(calls.jev[0].headers.Authorization, "Bearer or-key");

  const n1 = calls.jev.find(
    (c) => c.request.state.transaction.bank_description === "SAFEWAY #5678 SAN FRANCISCO CA"
  ).request;
  assert.strictEqual(n1.model, "typesafe/jev-1.13");
  assert.deepStrictEqual(n1.state, {
    transaction: {
      bank_description: "SAFEWAY #5678 SAN FRANCISCO CA",
      amount: -18.5,
      date: "2026-10-01",
    },
  });
  assert.strictEqual(n1.questions.precedent.type, "choice");
  assert.deepStrictEqual(n1.questions.precedent.criteria, {
    p1: {
      bank_description: "SAFEWAY #1234 SAN FRANCISCO CA",
      category: "Groceries",
      amount: -42.1,
      date: iso(HISTORY[0][0]),
    },
    none: "No previous transaction is the same merchant or the same recurring payment as this one",
  });
});

test("a confident Jev pick copies the description and category and skips Gemini", () => {
  const { calls, written } = run({
    openRouterKey: "or-key",
    jev: () => jevAnswer("p1", 0.1),
    gemini: () =>
      geminiResponse([
        { transaction_id: "N3", updated_description: "Blue Bottle Coffee", category: "Restaurants" },
      ]),
  });

  const payload = JSON.parse(calls.gemini[0].request.contents[0].parts[0].text);
  assert.deepStrictEqual(payload.transactions.map((t) => t.transaction_id), ["N3"]);

  assert.deepStrictEqual(written.N1, {
    Category: "Groceries",
    Description: "Safeway",
    "AI AutoCat": "TRUE",
  });
  assert.strictEqual(written.N2.Category, "Utilities");
  assert.strictEqual(written.N2.Description, "PG&E");
  assert.strictEqual(written.N3.Category, "Restaurants");
});

test("Gemini is not called at all when Jev settles every row", () => {
  const { calls, written } = run({
    openRouterKey: "or-key",
    newRows: [NEW_ROWS[0], NEW_ROWS[1]],
    jev: () => jevAnswer("p1", 0.2),
    gemini: () => {
      throw new Error("Gemini should not be called");
    },
  });
  assert.strictEqual(calls.gemini.length, 0);
  assert.strictEqual(written.N1.Category, "Groceries");
  assert.strictEqual(written.N2.Category, "Utilities");
});

test("Jev picks are gated on 1 - P(none) >= 0.5, not on confidence", () => {
  const { calls } = run({
    openRouterKey: "or-key",
    newRows: [NEW_ROWS[0]],
    jev: () => jevAnswer("p1", 0.6),
    gemini: () => geminiResponse([]),
  });
  const payload = JSON.parse(calls.gemini[0].request.contents[0].parts[0].text);
  assert.deepStrictEqual(payload.transactions.map((t) => t.transaction_id), ["N1"]);
});

test("none, an unknown option, and a failed request all fall through to Gemini", () => {
  let n = 0;
  const answers = [
    jevAnswer("none", 0.9),
    jevAnswer("p7", 0.0),
  ];
  const { calls, logs } = run({
    openRouterKey: "or-key",
    newRows: [NEW_ROWS[0], NEW_ROWS[1], ["N4", "SAFEWAY #9999 SAN FRANCISCO CA", -3, "Date(2026,9,4)"]],
    jev: () => answers[n++] || response(500, { error: { message: "boom" } }),
    gemini: () => geminiResponse([]),
  });
  const payload = JSON.parse(calls.gemini[0].request.contents[0].parts[0].text);
  assert.deepStrictEqual(payload.transactions.map((t) => t.transaction_id), ["N1", "N2", "N4"]);
  assert.ok(logs.some((l) => typeof l === "object" && l.jevFailed === 1 && l.jevTaken === 0));
});

test("a Jev pick of a To Be Categorized precedent goes to Gemini instead", () => {
  const { calls } = run({
    openRouterKey: "or-key",
    newRows: [["N5", "ZXQ MYSTERY CHARGE", -9, "Date(2026,9,5)"]],
    jev: () => jevAnswer("p1", 0.0),
    gemini: () => geminiResponse([]),
  });
  assert.strictEqual(calls.jev.length, 1);
  const payload = JSON.parse(calls.gemini[0].request.contents[0].parts[0].text);
  assert.deepStrictEqual(payload.transactions.map((t) => t.transaction_id), ["N5"]);
});

test("Jev picks are still written when the Gemini call fails", () => {
  const { written } = run({
    openRouterKey: "or-key",
    jev: () => jevAnswer("p1", 0.1),
    gemini: () => response(500, { error: { message: "unavailable" } }),
  });
  assert.strictEqual(written.N1.Category, "Groceries");
  assert.strictEqual(written.N2.Category, "Utilities");
  assert.strictEqual(written.N3, undefined);
});

test("Jev requests go out at most 32 at a time", () => {
  const many = [];
  for (let i = 0; i < 70; i++) {
    many.push(["M" + i, "SAFEWAY #" + (2000 + i) + " SAN FRANCISCO CA", -1, "Date(2026,9,1)"]);
  }
  const batchSizes = [];
  const rows = HISTORY.map((r) => r.slice()).concat(
    many.map((n) => {
      const row = new Array(HEADERS.length).fill("");
      row[HEADERS.indexOf("Transaction ID")] = n[0];
      row[HEADERS.indexOf("Full Description")] = n[1];
      return row;
    })
  );
  const spreadsheet = new FakeSpreadsheet({
    Transactions: new FakeSheet("Transactions", HEADERS, rows),
    Categories: new FakeSheet("Categories", ["Category", "Group"], CATEGORIES),
  });
  const { context } = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: { GCP_PROJECT_ID: "proj", OPENROUTER_API_KEY: "k" },
    fetch: (url) => ({ getContentText: () => gvizResponse(many) }),
    fetchAll: (requests) => {
      batchSizes.push(requests.length);
      return requests.map(() => jevAnswer("p1", 0.0));
    },
  });
  context.categorizeUncategorizedTransactions();
  assert.deepStrictEqual(batchSizes, [32, 32, 6]);
});

// --- Choosing previous transactions by amount -------------------------------

function hit(id, score, amount, category) {
  return { id, score, amount, category: category || "C" };
}

test("without an exact tie beyond six slots, the top six are taken in rank order", () => {
  const { context } = loadScripts(SOURCES);
  const hits = [
    hit("a", 2.0, -5), hit("b", 1.9, -500), hit("c", 1.8, -5), hit("d", 1.7, -5),
    hit("e", 1.6, -5), hit("f", 1.5, -5), hit("g", 1.4, -5),
  ];
  const chosen = plain(context.closestByAmount(hits, -500, 6)).map((h) => h.id);
  assert.deepStrictEqual(chosen, ["a", "b", "c", "d", "e", "f"]);
});

test("six or fewer exact ties keep rank order, even if amounts differ", () => {
  const { context } = loadScripts(SOURCES);
  const hits = [hit("a", 1, -5), hit("b", 1, -2850), hit("c", 0.5, -2850)];
  const chosen = plain(context.closestByAmount(hits, -2850, 6)).map((h) => h.id);
  assert.deepStrictEqual(chosen, ["a", "b", "c"]);
});

test("near ties do not count; only scores within 0.9999 of the top", () => {
  const { context } = loadScripts(SOURCES);
  const hits = [];
  for (let i = 0; i < 8; i++) hits.push(hit("t" + i, 1 - i * 0.01, i === 7 ? -2850 : -5));
  const chosen = plain(context.closestByAmount(hits, -2850, 6)).map((h) => h.id);
  assert.deepStrictEqual(chosen, ["t0", "t1", "t2", "t3", "t4", "t5"]);
});

test("more than six exact ties: closest in amount, one per category first, opposite signs last", () => {
  const { context } = loadScripts(SOURCES);
  const hits = [
    hit("refund", 1, 2850, "Rent"),        // opposite sign: distance 1000
    hit("small1", 1, -20, "Gifts"),
    hit("small2", 1, -25, "Gifts"),
    hit("small3", 1, -30, "Gifts"),
    hit("rent1", 1, -2850, "Rent"),
    hit("rent2", 1, -2850, "Rent"),
    hit("rent3", 1, -2850, "Rent"),
    hit("big", 1, -1000, "Home"),
    hit("lower", 0.5, -2850, "Rent"),      // not tied, never considered
  ];
  const chosen = plain(context.closestByAmount(hits, -2850, 6)).map((h) => h.id);
  // Distance order: rent1-3 (0), big, small3, small2, small1, refund.
  // One per category first: rent1 (Rent), big (Home), small3 (Gifts).
  // Fill from the same order: rent2, rent3, small2. Returned in distance order.
  assert.deepStrictEqual(chosen, ["rent1", "rent2", "rent3", "big", "small3", "small2"]);
});

test("a rent check is shown its rent precedents, not the newest random checks", () => {
  // Every CHECK #NNNN row scores the same against another check, so word
  // overlap alone ranks by date and the rent checks fall off the list.
  const headers = HEADERS;
  const rows = [];
  for (let i = 0; i < 10; i++) {
    rows.push([daysAgo(5 + i), "Check", "Gifts", -40 - i, "CHECK #" + (5000 + i), "R" + i, ""]);
  }
  for (let m = 0; m < 3; m++) {
    rows.push([daysAgo(40 + 30 * m), "Rent", "Rent", -2850, "CHECK #" + (4000 + m), "RENT" + m, ""]);
  }
  // Other merchants, so "check" is not in every row and still scores.
  for (let i = 0; i < 40; i++) {
    const d = ["SAFEWAY", "NETFLIX", "PG&E", "AMAZON"][i % 4] + " PAYMENT " + i;
    rows.push([daysAgo(3 + i), d, "Gifts", -10, d, "O" + i, ""]);
  }
  rows.push(["", "", "", "", "CHECK #4100", "N1", ""]);
  const spreadsheet = new FakeSpreadsheet({
    Transactions: new FakeSheet("Transactions", headers, rows),
    Categories: new FakeSheet("Categories", ["Category", "Group"], [
      ["Rent", "Home"], ["Gifts", "Other"],
    ]),
  });
  const geminiCalls = [];
  const { context } = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: { GCP_PROJECT_ID: "proj" },
    fetch: (url, params) => {
      if (url.includes("/gviz/")) {
        return { getContentText: () => gvizResponse([["N1", "CHECK #4100", -2850, "Date(2026,9,2)"]]) };
      }
      geminiCalls.push(JSON.parse(params.payload));
      return geminiResponse([]);
    },
  });
  context.categorizeUncategorizedTransactions();

  const payload = JSON.parse(geminiCalls[0].contents[0].parts[0].text);
  const previous = payload.transactions[0].previous_transactions;
  assert.strictEqual(previous.length, 6);
  assert.deepStrictEqual(
    previous.slice(0, 3).map((p) => [p.category, p.amount]),
    [["Rent", -2850], ["Rent", -2850], ["Rent", -2850]]
  );
});

// --- Choosing previous transactions by amount, end to end through the search ---
// Real search results fed to closestByAmount with the production settings.

function searchByAmount(context, docs, query, amount) {
  const searcher = new context.TFIDFSearch(
    docs.map((d) => ({
      id: d.id,
      text: d.rawText,
      updatedText: "",
      category: d.categoryId,
      amount: d.amount,
      date: d.date ? new Date(d.date) : null,
    })),
    { useStopWords: true, matchThreshold: 0.25, minTermSize: 3 }
  );
  return plain(
    context.closestByAmount(
      searcher.search(query, context.PRECEDENT_CANDIDATES),
      amount,
      context.EXAMPLES_PER_ROW
    )
  );
}

const UNRELATED = [
  "ELECTRIC UTILITY", "WATER DISTRICT", "BLUE BOTTLE COFFEE", "CHEVRON STATION",
  "NETFLIX STREAMING", "PACIFIC TELEPHONE", "GOLDEN GATE PARKING", "WHOLE FOODS",
  "AIRLINE TICKETS", "HARDWARE STORE",
].map((rawText, i) => ({ id: "unrelated-" + i, rawText, categoryId: "other", amount: -20 }));

test("shows the matches closest in amount when more tie on the text than there are slots", () => {
  const { context } = loadScripts(SOURCES);
  // const declarations are not properties of the vm context; expose them.
  context.PRECEDENT_CANDIDATES = vmConst(context, "PRECEDENT_CANDIDATES");
  context.EXAMPLES_PER_ROW = vmConst(context, "EXAMPLES_PER_ROW");
  const checks = [-40, -120, -310, -455, -620, -880, -2850, -1020].map((value, i) => ({
    id: "check-" + i,
    rawText: "CHECK #" + (1001 + i),
    categoryId: value === -2850 ? "rent" : value === -40 ? "gift" : "misc",
    amount: value,
    date: "2026-0" + (1 + i) + "-02",
  }));
  const chosen = searchByAmount(context, [...checks, ...UNRELATED], "CHECK #1060", -2850);
  assert.deepStrictEqual(chosen.map((x) => x.amount), [-2850, -1020, -880, -620, -455, -40]);
});

test("keeps the text ranking when the matches do not tie", () => {
  const { context } = loadScripts(SOURCES);
  context.PRECEDENT_CANDIDATES = vmConst(context, "PRECEDENT_CANDIDATES");
  context.EXAMPLES_PER_ROW = vmConst(context, "EXAMPLES_PER_ROW");
  const chosen = searchByAmount(
    context,
    [
      ...UNRELATED,
      { id: "store", rawText: "SAFEWAY MARKET", categoryId: "groceries", amount: -900 },
      { id: "fuel", rawText: "SAFEWAY FUEL", categoryId: "fuel", amount: -10 },
    ],
    "SAFEWAY MARKET 123",
    -10
  );
  assert.deepStrictEqual(chosen.map((x) => x.category), ["groceries", "fuel"]);
});

function vmConst(context, name) {
  return require("vm").runInContext(name, context);
}

test("each Jev answer is logged with its choice, match probability and confidence", () => {
  const { logs } = run({
    openRouterKey: "or-key",
    newRows: [NEW_ROWS[0], NEW_ROWS[1]],
    jev: (req) =>
      req.state.transaction.bank_description.startsWith("SAFEWAY")
        ? jevAnswer("p1", 0.1)
        : jevAnswer("none", 0.8),
    gemini: () => geminiResponse([]),
  });
  const at = logs.indexOf("Jev decisions:");
  assert.ok(at !== -1);
  assert.deepStrictEqual(plain(logs[at + 1]), [
    {
      transaction_id: "N1",
      description: "SAFEWAY #5678 SAN FRANCISCO CA",
      choice: "p1 (Groceries)",
      match: 0.9,
      confidence: 0.9,
      taken: true,
    },
    {
      transaction_id: "N2",
      description: "PG&E WEB ONLINE PAYMENT 1111111",
      choice: "none",
      match: 0.2,
      confidence: 0.9,
      taken: false,
    },
  ]);
});

// --- Writing after each stage -----------------------------------------------

test("Jev's matches are written before Gemini is called, then Gemini's in a second sweep", () => {
  const { logs, written } = run({
    openRouterKey: "or-key",
    jev: () => jevAnswer("p1", 0.1),
    gemini: () =>
      geminiResponse([
        { transaction_id: "N3", updated_description: "Blue Bottle Coffee", category: "Restaurants" },
      ]),
  });
  assert.ok(written.N1 && written.N2 && written.N3);

  const order = logs.filter((l) => typeof l === "string");
  const jevWrite = order.indexOf("Writing 2 Jev match(es) into your sheet...");
  const geminiAsk = order.findIndex((l) => l.startsWith("Using Gemini"));
  const geminiWrite = order.indexOf("Writing 1 Gemini answer(s) into your sheet...");
  assert.ok(jevWrite !== -1 && jevWrite < geminiAsk && geminiAsk < geminiWrite);
  const rows = HISTORY.length + NEW_ROWS.length;
  assert.deepStrictEqual(
    order.filter((l) => l.startsWith("Found ") && l.includes(" of ")),
    [
      "Found 2 of 2 transactions in first " + rows + " rows.",
      "Found 1 of 1 transactions in first " + rows + " rows.",
    ]
  );
});


// --- Assigning Transaction IDs to rows that have none -----------------------

test("ulid() is 26 Crockford base-32 characters that sort by time", () => {
  const { context } = loadScripts(SOURCES);
  const a = context.ulid(1469918176385);
  const b = context.ulid(1469918176386);
  assert.match(a, /^[0-9A-HJKMNP-TV-Z]{26}$/);
  // The timestamp example from the ULID spec (github.com/ulid/spec).
  assert.strictEqual(a.slice(0, 10), "01ARYZ6S41");
  assert.ok(a.slice(0, 10) < b.slice(0, 10));
});

/**
 * A sheet whose gviz query is answered from the sheet's current contents, so
 * ids written before the query are visible to it, as on a real sheet.
 */
function liveRun(sheetRows, geminiAnswer, runOptions, extra) {
  const sheet = new FakeSheet("Transactions", HEADERS, sheetRows);
  const spreadsheet = new FakeSpreadsheet({
    Transactions: sheet,
    Categories: new FakeSheet("Categories", ["Category", "Group"], CATEGORIES),
  });
  let flushed = false;
  const geminiCalls = [];
  const geminiGroups = [];
  const col = (name) => HEADERS.indexOf(name);
  const loaded = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: { GCP_PROJECT_ID: "proj" },
    onFlush: () => (flushed = true),
    fetchAll: (requests) => {
      geminiGroups.push(requests.length);
      return requests.map((r) => loaded.sandbox.UrlFetchApp.fetch(r.url, r));
    },
    fetch: (url, params) => {
      if (url.includes("/gviz/")) {
        const idWrites = sheet.writes.filter((w) => w.column === "Transaction ID");
        assert.ok(flushed || idWrites.length === 0, "ids must be flushed before the gviz query");
        const rows = sheet.data
          .slice(1)
          .filter((r) => r[col("Full Description")] !== "" && r[col("Category")] === "")
          .slice(0, B)
          .map((r) => [
            r[col("Transaction ID")] === "" ? null : r[col("Transaction ID")],
            r[col("Full Description")],
            r[col("Amount")],
            "Date(2023,8,25)",
          ]);
        return { getContentText: () => gvizResponse(rows) };
      }
      const payload = JSON.parse(JSON.parse(params.payload).contents[0].parts[0].text);
      geminiCalls.push(payload.transactions.map((t) => t.transaction_id));
      if (extra && extra.geminiFails && extra.geminiFails(geminiCalls.length)) {
        return response(500, { error: { message: "down" } });
      }
      return geminiResponse(payload.transactions.map(geminiAnswer));
    },
  });
  loaded.context.categorizeUncategorizedTransactions(runOptions);
  return { sheet, logs: loaded.logs, geminiCalls, geminiGroups };
}

test("a row with no Transaction ID is given one, sent with it, and written back", () => {
  const rows = HISTORY.map((r) => r.slice()).concat([
    [new Date(2023, 8, 25), "", "", -19.99, "Krispy Kreme", "", ""],
  ]);
  const sent = [];
  const { sheet, logs } = liveRun(rows, (t) => {
    sent.push(t.transaction_id);
    return { transaction_id: t.transaction_id, updated_description: "Krispy Kreme", category: "Restaurants" };
  });
  const last = sheet.data[sheet.data.length - 1];
  assert.match(last[HEADERS.indexOf("Transaction ID")], /^autocat:[0-9A-HJKMNP-TV-Z]{26}$/);
  assert.deepStrictEqual(sent, [last[HEADERS.indexOf("Transaction ID")]]);
  assert.strictEqual(last[HEADERS.indexOf("Category")], "Restaurants");
  assert.ok(logs.includes("Assigned a Transaction ID to 1 row(s) that had none: rows " + sheet.data.length));
});

test("only rows about to be sent get an id; categorized rows and rows with ids are left alone", () => {
  const rows = HISTORY.map((r) => r.slice()).concat([
    [new Date(2023, 8, 25), "Done", "Restaurants", -1, "CATEGORIZED NO ID", "", ""], // has a category
    [new Date(2023, 8, 25), "", "", -1, "", "", ""], // no Full Description
    [new Date(2023, 8, 25), "", "", -1, "HAS AN ID", "EXISTING", ""],
    [new Date(2023, 8, 25), "", "", -1, "NEEDS AN ID", "", ""],
  ]);
  const { sheet } = liveRun(rows, (t) => ({
    transaction_id: t.transaction_id,
    updated_description: "x",
    category: "Restaurants",
  }));
  const idWrites = sheet.writes.filter((w) => w.column === "Transaction ID");
  assert.strictEqual(idWrites.length, 1);
  assert.strictEqual(idWrites[0].row, sheet.data.length);
  const n = sheet.data.length;
  assert.strictEqual(sheet.data[n - 4][HEADERS.indexOf("Transaction ID")], "");
  assert.strictEqual(sheet.data[n - 3][HEADERS.indexOf("Transaction ID")], "");
  assert.strictEqual(sheet.data[n - 2][HEADERS.indexOf("Transaction ID")], "EXISTING");
});

test("a row past the first batch gets its id only when its own batch starts", () => {
  // A full batch of uncategorized rows that already have ids comes first, so
  // the id-less row after them is not in the first batch; it is given an id
  // for the second.
  const rows = HISTORY.map((r) => r.slice());
  for (let i = 0; i < B; i++) {
    rows.push([new Date(2023, 8, 25), "", "", -1, "ROW " + i, "ID" + i, ""]);
  }
  rows.push([new Date(2023, 8, 25), "", "", -1, "LATER NO ID", "", ""]);
  const { sheet, logs, geminiCalls } = liveRun(rows, (t) => ({
    transaction_id: t.transaction_id,
    updated_description: "x",
    category: "to-be-categorized",
  }));
  const strings = logs.filter((l) => typeof l === "string");
  const batch1Write = strings.findIndex((l) => /^Writing \d+ Gemini answer\(s\)/.test(l));
  const assigned = strings.findIndex((l) => l.startsWith("Assigned a Transaction ID to 1 row(s)"));
  const batch2 = strings.indexOf("--- Batch 2 ---");
  assert.ok(batch1Write < assigned && assigned < batch2);
  assert.strictEqual(geminiCalls.length, requestsFor(B) + 1);
  assert.deepStrictEqual(geminiCalls[geminiCalls.length - 1].length, 1);
  assert.match(sheet.data[sheet.data.length - 1][HEADERS.indexOf("Transaction ID")], /^autocat:/);
});

// --- Looping over batches ---------------------------------------------------

function manyRows(n) {
  const rows = HISTORY.map((r) => r.slice());
  for (let i = 0; i < n; i++) {
    rows.push([new Date(2023, 8, 25), "", "", -1, "ROW " + i, "ID" + i, ""]);
  }
  return rows;
}

const answerAll = (t) => ({
  transaction_id: t.transaction_id,
  updated_description: "x",
  category: "Restaurants",
});

test("a batch's Gemini rows go 50 to a request, three requests at a time", () => {
  assert.deepStrictEqual([SETTINGS.perRequest, SETTINGS.inFlight], [50, 3]);
  const rows = B - 20; // a partial last request
  const { geminiCalls, geminiGroups } = liveRun(manyRows(rows), answerAll);
  const n = requestsFor(rows);
  assert.deepStrictEqual(
    geminiCalls.map((c) => c.length),
    Array.from({ length: n }, (_, i) => (i < n - 1 ? 50 : rows - 50 * (n - 1)))
  );
  const groups = [];
  for (let left = n; left > 0; left -= 3) groups.push(Math.min(3, left));
  assert.deepStrictEqual(geminiGroups, groups);
});

test("a run works through several batches until nothing is left", () => {
  const total = 2 * B + 100;
  const { sheet, geminiCalls, logs } = liveRun(manyRows(total), answerAll);
  assert.strictEqual(geminiCalls.length, 2 * requestsFor(B) + requestsFor(100));
  assert.ok(geminiCalls.every((c) => c.length <= 50));
  const cat = HEADERS.indexOf("Category");
  assert.strictEqual(sheet.data.slice(1).filter((r) => r[cat] === "").length, 0);
  assert.ok(logs.some((l) => typeof l === "object" && l.batches === 3 && l.answersWritten === total));
});

test("a run stops starting batches when the next would overrun the time budget", () => {
  // Each full batch takes 100s on this clock; with a 250s budget, after two
  // batches (200s) another would end at 300s, so the run stops.
  let clock = 0;
  const now = () => clock;
  const { logs } = liveRun(
    manyRows(4 * B),
    (t) => {
      clock += 100000 / B;
      return answerAll(t);
    },
    { budgetMs: 250000, now }
  );
  assert.strictEqual(logs.filter((l) => typeof l === "string" && l.startsWith("--- Batch")).length, 2);
  assert.ok(logs.includes("Stopping before batch 3 to stay within 250s; run again for the rest."));
});

test("the default budget is five minutes, and a trigger's event object is ignored", () => {
  const { context } = loadScripts(SOURCES);
  assert.strictEqual(require("vm").runInContext("RUN_TIME_BUDGET_MS", context), 300000);
  // A time-driven trigger passes an event object; the run must still go.
  const { geminiCalls } = liveRun(manyRows(3), answerAll, { authMode: "FULL", triggerUid: "1" });
  assert.strictEqual(geminiCalls.length, 1);
});

test("when one Gemini request fails, the others are written and the run stops", () => {
  // The second of three requests fails.
  const { sheet, geminiCalls, logs } = liveRun(manyRows(B + 300), answerAll, undefined, {
    geminiFails: (n) => n === 2,
  });
  assert.strictEqual(geminiCalls.length, requestsFor(B)); // batch 1 only
  const cat = HEADERS.indexOf("Category");
  const full = HEADERS.indexOf("Full Description");
  const written = sheet.data
    .slice(1)
    .filter((r) => String(r[full]).startsWith("ROW ") && r[cat] === "Restaurants").length;
  assert.strictEqual(written, B - 50);
  assert.ok(logs.includes("1 of " + requestsFor(B) + " Gemini request(s) failed."));
  assert.ok(logs.includes("Stopping: the Gemini call failed."));
});

test("rows that cannot be written are not sent again in the same run", () => {
  // Gemini answers with ids that are not on the sheet, so nothing is written
  // and the same rows come back from the query.
  const { geminiCalls, logs } = liveRun(manyRows(B + 50), (t) => ({
    transaction_id: "MISSING-" + t.transaction_id,
    updated_description: "x",
    category: "Restaurants",
  }));
  assert.strictEqual(geminiCalls.length, requestsFor(B)); // the first batch, once
  assert.ok(
    logs.includes(
      B + " uncategorized transaction(s) were already tried in this run and could not be written; stopping."
    )
  );
});

test("Jev's matches are written after each group of parallel requests", () => {
  // 70 rows with a previous transaction: Jev groups of 32, 32 and 6, each
  // written before the next group is sent.
  const many = [];
  for (let i = 0; i < 70; i++) {
    many.push(["M" + i, "SAFEWAY #" + (2000 + i) + " SAN FRANCISCO CA", -1, "Date(2026,9,1)"]);
  }
  const rows = HISTORY.map((r) => r.slice()).concat(
    many.map((n) => {
      const row = new Array(HEADERS.length).fill("");
      row[HEADERS.indexOf("Transaction ID")] = n[0];
      row[HEADERS.indexOf("Full Description")] = n[1];
      return row;
    })
  );
  const sheet = new FakeSheet("Transactions", HEADERS, rows);
  const spreadsheet = new FakeSpreadsheet({
    Transactions: sheet,
    Categories: new FakeSheet("Categories", ["Category", "Group"], CATEGORIES),
  });
  const writtenBeforeGroup = [];
  const { context, logs } = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: { GCP_PROJECT_ID: "proj", OPENROUTER_API_KEY: "k" },
    fetch: () => ({ getContentText: () => gvizResponse(many) }),
    fetchAll: (requests) => {
      writtenBeforeGroup.push(new Set(sheet.writes.map((w) => w.row)).size);
      return requests.map(() => jevAnswer("p1", 0.0));
    },
  });
  context.categorizeUncategorizedTransactions();
  assert.deepStrictEqual(writtenBeforeGroup, [0, 32, 64]);
  assert.deepStrictEqual(
    logs.filter((l) => typeof l === "string" && l.includes("Jev match(es)")),
    [
      "Writing 32 Jev match(es) into your sheet...",
      "Writing 32 Jev match(es) into your sheet...",
      "Writing 6 Jev match(es) into your sheet...",
    ]
  );
});

test("Gemini's answers are written after each group of parallel requests", () => {
  // A full batch goes to Gemini in requests of 50, three at a time; each
  // group's answers are in the sheet before the next group is sent.
  const rows = manyRows(330);
  const sheet = new FakeSheet("Transactions", HEADERS, rows);
  const spreadsheet = new FakeSpreadsheet({
    Transactions: sheet,
    Categories: new FakeSheet("Categories", ["Category", "Group"], CATEGORIES),
  });
  const col = (name) => HEADERS.indexOf(name);
  const pending = () =>
    sheet.data
      .slice(1)
      .filter((r) => r[col("Full Description")] !== "" && r[col("Category")] === "");
  const writtenBeforeGroup = [];
  let firstPending = null;
  const { context } = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: { GCP_PROJECT_ID: "proj" },
    fetch: (url) => {
      const batch = pending().slice(0, B).map((r) => [r[col("Transaction ID")], r[col("Full Description")], -1, "Date(2023,8,25)"]);
      return { getContentText: () => gvizResponse(batch) };
    },
    fetchAll: (requests) => {
      if (firstPending === null) firstPending = pending().length;
      writtenBeforeGroup.push(firstPending - pending().length);
      return requests.map((r) => {
        const payload = JSON.parse(JSON.parse(r.payload).contents[0].parts[0].text);
        return geminiResponse(payload.transactions.map(answerAll));
      });
    },
  });
  context.categorizeUncategorizedTransactions();
  // Batch 1 is B rows: groups of 3, 2 requests (150, 100 rows).
  assert.deepStrictEqual(writtenBeforeGroup.slice(0, 2), [0, 150]);
});

// --- Timing and the missing-id scan -----------------------------------------

test("each batch and the run log where their time went", () => {
  // Gemini takes 4s per request on this clock; nothing else takes time.
  let clock = 0;
  const { logs } = liveRun(
    manyRows(200),
    (t) => {
      clock += 4000 / 50;
      return answerAll(t);
    },
    { budgetMs: 300000, now: () => clock }
  );
  // 200 rows: a group of three 50-row requests (12s), written, then one more (4s).
  const batch = logs.find((l) => typeof l === "string" && l.startsWith("Batch 1 timing"));
  assert.strictEqual(batch, "Batch 1 timing (200 rows): 16.0s total: Gemini requests 16.0s");
  const run = logs.find((l) => typeof l === "string" && l.startsWith("Run timing"));
  assert.strictEqual(
    run,
    "Run timing: 16.0s total: Gemini requests 16.0s; first answer written after 12.0s"
  );
});

test("the sheet is scanned for missing ids only when the query returns a row without one", () => {
  const rows = manyRows(5);
  let scans = 0;
  const sheet = new FakeSheet("Transactions", HEADERS, rows);
  const spreadsheet = new FakeSpreadsheet({
    Transactions: sheet,
    Categories: new FakeSheet("Categories", ["Category", "Group"], CATEGORIES),
  });
  const col = (name) => HEADERS.indexOf(name);
  const { context } = loadScripts(SOURCES, {
    spreadsheet,
    scriptProperties: { GCP_PROJECT_ID: "proj" },
    fetch: (url, params) => {
      if (url.includes("/gviz/")) {
        const batch = sheet.data
          .slice(1)
          .filter((r) => r[col("Full Description")] !== "" && r[col("Category")] === "")
          .map((r) => [r[col("Transaction ID")] || null, r[col("Full Description")], -1, "Date(2023,8,25)"]);
        return { getContentText: () => gvizResponse(batch) };
      }
      const payload = JSON.parse(JSON.parse(params.payload).contents[0].parts[0].text);
      return geminiResponse(payload.transactions.map(answerAll));
    },
  });
  const original = context.assignMissingTransactionIds;
  context.assignMissingTransactionIds = function () {
    scans++;
    return original();
  };
  context.categorizeUncategorizedTransactions();
  assert.strictEqual(scans, 0, "every row had an id, so no scan");

  // Now a row without an id: one scan, and it is written like the others.
  sheet.data.push([new Date(2023, 8, 25), "", "", -1, "NO ID ROW", "", ""]);
  context.categorizeUncategorizedTransactions();
  assert.strictEqual(scans, 1);
  const last = sheet.data[sheet.data.length - 1];
  assert.match(last[col("Transaction ID")], /^autocat:/);
  assert.strictEqual(last[col("Category")], "Restaurants");
});
