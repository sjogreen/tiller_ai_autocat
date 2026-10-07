/**
 * Test harness: loads the .gs files into a sandboxed context with stub
 * implementations of the Apps Script globals they depend on.
 *
 * These tests exist to prove the performance refactor did not change behaviour,
 * so several stubs deliberately reproduce Apps Script quirks (notably that
 * getRange("12") is invalid A1 notation and throws).
 */
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { execFileSync } = require("child_process");

const REPO_ROOT = path.resolve(__dirname, "..");

function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1; // zero based
}

function columnIndexToLetter(index) {
  let letter = "";
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

/**
 * A minimal spreadsheet stub that records every write, so tests can assert
 * exactly which cells were touched.
 */
class FakeSheet {
  constructor(name, headers, rows, maxRows) {
    this.name = name;
    this.headers = headers;
    this.maxRows = maxRows || 1000;
    this.data = [headers.slice()];
    rows.forEach((r) => this.data.push(r.slice()));
    this.writes = []; // { a1, value, via }

    // Cells whose value comes from a sheet-level ARRAYFORMULA. In real Sheets,
    // writing any literal into one of these replaces the spilled formula and
    // breaks it for the whole column - even writing back the value just read.
    this.formulaCells = new Set();
    this.brokenFormulas = [];
  }

  /** Marks cells as ARRAYFORMULA-derived, by A1 notation. */
  markFormulaDerived(a1List) {
    a1List.forEach((a1) => this.formulaCells.add(a1));
    return this;
  }

  _recordWrite(a1, rowIndex, colIndex, value, via, cellsInWrite) {
    if (this.formulaCells.has(a1)) {
      this.brokenFormulas.push({ a1, value, via });
    }
    this.writes.push({
      a1,
      row: rowIndex + 1,
      column: this.headers[colIndex],
      value,
      via,
      cellsInWrite,
    });
  }

  getLastRow() {
    return this.data.length;
  }

  getLastColumn() {
    return this.headers.length;
  }

  _cell(rowIndex, colIndex) {
    const row = this.data[rowIndex];
    if (!row) return "";
    const value = row[colIndex];
    return value === undefined ? "" : value;
  }

  _valuesFor(startRow, startCol, numRows, numCols) {
    const out = [];
    for (let r = 0; r < numRows; r++) {
      const row = [];
      for (let c = 0; c < numCols; c++) {
        row.push(this._cell(startRow + r, startCol + c));
      }
      out.push(row);
    }
    return out;
  }

  getRange(a1OrRow, col, numRows, numCols) {
    if (typeof a1OrRow === "number") {
      return new FakeRange(
        this,
        a1OrRow - 1,
        col - 1,
        numRows === undefined ? 1 : numRows,
        numCols === undefined ? 1 : numCols
      );
    }
    return this._rangeFromA1(String(a1OrRow));
  }

  _rangeFromA1(a1) {
    let m;

    // Whole rows, e.g. "1:1"
    if ((m = /^(\d+):(\d+)$/.exec(a1))) {
      const start = parseInt(m[1], 10) - 1;
      const end = parseInt(m[2], 10) - 1;
      return new FakeRange(this, start, 0, end - start + 1, this.headers.length);
    }

    // Whole columns, e.g. "E:E"
    if ((m = /^([A-Z]+):([A-Z]+)$/.exec(a1))) {
      const start = columnLetterToIndex(m[1]);
      const end = columnLetterToIndex(m[2]);
      return new FakeRange(this, 0, start, this.maxRows, end - start + 1);
    }

    // Open ended column range, e.g. "B2:B" - runs to the bottom of the sheet,
    // which is why blank rows come back from getAllowedCategories.
    if ((m = /^([A-Z]+)(\d+):([A-Z]+)$/.exec(a1))) {
      const startCol = columnLetterToIndex(m[1]);
      const endCol = columnLetterToIndex(m[3]);
      const startRow = parseInt(m[2], 10) - 1;
      return new FakeRange(
        this,
        startRow,
        startCol,
        this.maxRows - startRow,
        endCol - startCol + 1
      );
    }

    // Bounded range, e.g. "B2:D9"
    if ((m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a1))) {
      const startCol = columnLetterToIndex(m[1]);
      const startRow = parseInt(m[2], 10) - 1;
      const endCol = columnLetterToIndex(m[3]);
      const endRow = parseInt(m[4], 10) - 1;
      return new FakeRange(
        this,
        startRow,
        startCol,
        endRow - startRow + 1,
        endCol - startCol + 1
      );
    }

    // Single cell, e.g. "C12"
    if ((m = /^([A-Z]+)(\d+)$/.exec(a1))) {
      return new FakeRange(
        this,
        parseInt(m[2], 10) - 1,
        columnLetterToIndex(m[1]),
        1,
        1
      );
    }

    // Anything else is invalid A1 notation. Apps Script throws here, and so must
    // we: a bare row number like "12" is what the old code built when an
    // optional column was missing.
    throw new Error('Invalid A1 notation: "' + a1 + '"');
  }

  getRangeList(a1List) {
    const ranges = a1List.map((a1) => this._rangeFromA1(a1));
    const sheet = this;
    return {
      setValue(value) {
        ranges.forEach((r) => r.setValue(value, "getRangeList"));
        return this;
      },
      getRanges() {
        return ranges;
      },
      _sheet: sheet,
    };
  }
}

class FakeRange {
  constructor(sheet, startRow, startCol, numRows, numCols) {
    this.sheet = sheet;
    this.startRow = startRow;
    this.startCol = startCol;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getA1Notation() {
    const a = columnIndexToLetter(this.startCol) + (this.startRow + 1);
    if (this.numRows === 1 && this.numCols === 1) return a;
    return (
      a +
      ":" +
      columnIndexToLetter(this.startCol + this.numCols - 1) +
      (this.startRow + this.numRows)
    );
  }

  getValues() {
    return this.sheet._valuesFor(
      this.startRow,
      this.startCol,
      this.numRows,
      this.numCols
    );
  }

  getRowIndex() {
    return this.startRow + 1;
  }

  setValue(value, via) {
    // Record every cell this write touches, so tests can assert that a write
    // never spills into a column the script does not own.
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        const rowIndex = this.startRow + r;
        const colIndex = this.startCol + c;
        while (this.sheet.data.length <= rowIndex) this.sheet.data.push([]);
        this.sheet.data[rowIndex][colIndex] = value;
        this.sheet._recordWrite(
          columnIndexToLetter(colIndex) + (rowIndex + 1),
          rowIndex,
          colIndex,
          value,
          via || "setValue",
          this.numRows * this.numCols
        );
      }
    }
    return this;
  }

  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const rowIndex = this.startRow + r;
        const colIndex = this.startCol + c;
        while (this.sheet.data.length <= rowIndex) this.sheet.data.push([]);
        this.sheet.data[rowIndex][colIndex] = values[r][c];
        this.sheet._recordWrite(
          columnIndexToLetter(colIndex) + (rowIndex + 1),
          rowIndex,
          colIndex,
          values[r][c],
          "setValues",
          values.length * values[r].length
        );
      }
    }
    return this;
  }

  // Mirrors Apps Script's default substring matching, first match wins.
  createTextFinder(query) {
    const range = this;
    return {
      findNext() {
        for (let r = 0; r < range.numRows; r++) {
          for (let c = 0; c < range.numCols; c++) {
            const value = range.sheet._cell(
              range.startRow + r,
              range.startCol + c
            );
            if (
              value !== "" &&
              String(value).indexOf(String(query)) !== -1
            ) {
              return new FakeRange(
                range.sheet,
                range.startRow + r,
                range.startCol + c,
                1,
                1
              );
            }
          }
        }
        return null;
      },
    };
  }
}

class FakeSpreadsheet {
  constructor(sheets) {
    this.sheets = sheets;
  }
  getId() {
    return "fake-spreadsheet-id";
  }
  getSheetByName(name) {
    return this.sheets[name] || null;
  }
  getActiveSheet() {
    return this.sheets[Object.keys(this.sheets)[0]];
  }
  toast() {}
  insertSheet(name) {
    this.sheets[name] = new FakeSheet(name, [], []);
    return this.sheets[name];
  }
}

/**
 * Answers a gviz query against a FakeSheet, for the query forms the script
 * uses: SELECT of column letters, WHERE of "X is null", "X is not null" and
 * "X >= date 'YYYY-MM-DD'" joined by AND, optional ORDER BY X asc|desc, and
 * LIMIT. Blank cells are null; dates come back as gviz's "Date(y,m,d)".
 */
function gvizFromSheet(sheet, query) {
  const m = /^SELECT (.+?) WHERE (.+?)(?: ORDER BY ([A-Z]+) (asc|desc))?(?: LIMIT (\d+))?$/.exec(query);
  if (!m) throw new Error("gviz emulator cannot read: " + query);
  const cols = m[1].split(",").map((c) => columnLetterToIndex(c.trim()));
  const isNull = (v) => v === "" || v === null || v === undefined;
  const day = (v) => {
    if (!(v instanceof Date)) return null;
    const pad = (n) => (n < 10 ? "0" : "") + n;
    return v.getFullYear() + "-" + pad(v.getMonth() + 1) + "-" + pad(v.getDate());
  };
  const conditions = m[2].split(" AND ").map((c) => {
    let x;
    if ((x = /^([A-Z]+) is not null$/.exec(c))) {
      const i = columnLetterToIndex(x[1]);
      return (row) => !isNull(row[i]);
    }
    if ((x = /^([A-Z]+) is null$/.exec(c))) {
      const i = columnLetterToIndex(x[1]);
      return (row) => isNull(row[i]);
    }
    if ((x = /^([A-Z]+) >= date '(\d{4}-\d{2}-\d{2})'$/.exec(c))) {
      const i = columnLetterToIndex(x[1]);
      return (row) => day(row[i]) !== null && day(row[i]) >= x[2];
    }
    throw new Error("gviz emulator cannot read condition: " + c);
  });
  let rows = sheet.data.slice(1).filter((row) => conditions.every((f) => f(row)));
  if (m[3]) {
    const i = columnLetterToIndex(m[3]);
    const dir = m[4] === "desc" ? -1 : 1;
    rows = rows
      .map((row, at) => ({ row, at }))
      .sort((a, b) => {
        const av = a.row[i] instanceof Date ? a.row[i].getTime() : -Infinity;
        const bv = b.row[i] instanceof Date ? b.row[i].getTime() : -Infinity;
        return av === bv ? a.at - b.at : (av - bv) * dir;
      })
      .map((x) => x.row);
  }
  if (m[5]) rows = rows.slice(0, Number(m[5]));
  const cell = (v) =>
    isNull(v)
      ? null
      : v instanceof Date
      ? { v: "Date(" + v.getFullYear() + "," + v.getMonth() + "," + v.getDate() + ")" }
      : { v };
  const table = {
    cols: cols.map((i) => ({ label: columnIndexToLetter(i) })),
    rows: rows.map((row) => ({ c: cols.map((i) => cell(row[i])) })),
  };
  return "/*O_o*/\ngoogle.visualization.Query.setResponse(" + JSON.stringify({ table }) + ");";
}

// The search-index query: it selects categorized rows and has no "is null"
// condition of its own. Every other gviz query is left to the test's stub.
function isIndexQuery(query) {
  return / is not null/.test(query) && !/(?<!not) is null/.test(query);
}

/**
 * Builds a context with the Apps Script globals stubbed, loads the given source
 * files into it, and returns the context plus the captured log.
 */
function loadScripts(sources, opts) {
  const options = opts || {};
  const logs = [];

  const sandbox = {
    console,
    Logger: {
      log(value) {
        logs.push(value);
      },
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => options.spreadsheet || null,
      flush: () => {
        if (options.onFlush) options.onFlush();
      },
      getUi: () => ({
        createMenu: () => ({ addItem: () => ({ addToUi: () => {} }) }),
      }),
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) =>
          options.scriptProperties && options.scriptProperties[key] !== undefined
            ? options.scriptProperties[key]
            : null,
      }),
    },
    ScriptApp: {
      getOAuthToken: () => "fake-oauth-token",
    },
    UrlFetchApp: {
      fetch: (url, params) => {
        // The search index's gviz query is answered from the fake sheet, so
        // every test builds its index from the sheet's own rows.
        const q = /[?&]tq=([^&]*)/.exec(url);
        const query = q ? decodeURIComponent(q[1]) : "";
        const sheet =
          options.spreadsheet && options.spreadsheet.getSheetByName("Transactions");
        if (url.includes("/gviz/") && sheet && isIndexQuery(query)) {
          if (options.onIndexQuery) options.onIndexQuery(query);
          const text = gvizFromSheet(sheet, query);
          return { getContentText: () => text };
        }
        if (options.fetch) return options.fetch(url, params);
        throw new Error("Unexpected UrlFetchApp.fetch call to " + url);
      },
      // Without a fetchAll stub, each request goes to the fetch stub, the same
      // as if it had been fetched on its own.
      fetchAll: (requests) =>
        options.fetchAll
          ? options.fetchAll(requests)
          : requests.map((r) => sandbox.UrlFetchApp.fetch(r.url, r)),
    },
    Utilities: {
      formatString: (fmt, ...args) => {
        let i = 0;
        return fmt.replace(/%s/g, () => String(args[i++]));
      },
    },
  };
  sandbox.globalThis = sandbox;

  const context = vm.createContext(sandbox);
  sources.forEach((src) => vm.runInContext(src, context, { filename: src.name }));

  return { context, sandbox, logs };
}

function readWorkingTree(file) {
  return fs.readFileSync(path.join(REPO_ROOT, file), "utf8");
}

/** Reads a file as it was at a given git revision, for before/after comparison. */
function readAtRevision(file, revision) {
  return execFileSync("git", ["show", (revision || "HEAD") + ":" + file], {
    cwd: REPO_ROOT,
    encoding: "utf8",
    maxBuffer: 10 * 1024 * 1024,
  });
}

module.exports = {
  gvizFromSheet,
  FakeSheet,
  FakeSpreadsheet,
  FakeRange,
  loadScripts,
  readWorkingTree,
  readAtRevision,
  columnIndexToLetter,
  columnLetterToIndex,
};
