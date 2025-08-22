// Diff two WorkbookModels (current vs baseline) and return per-sheet diffs and counts.

const CODE_NONE = 0;
const CODE_ADD = 1; // green
const CODE_REMOVE = 2; // red
// CODE_VALUE: calculated value difference (same non-empty formula)
// CODE_FORMULA: explicit change — formula text changed OR hardcoded literal changed
const CODE_VALUE = 3; // yellow (calculated value-only)
const CODE_FORMULA = 4; // orange (formula text or literal change)

function normFormula(f) {
  if (typeof f !== "string") return "";
  return f.trim().toUpperCase();
}

// Normalize text values coming from different sources (Office.js vs SheetJS) so
// visually identical strings compare equal.
// - Convert Windows/Mac line endings ("\r\n" or "\r") to "\n"
// - Convert non-breaking spaces (\u00A0) to regular spaces
// - Trim leading/trailing whitespace
function normTextForCompare(value) {
  // If it's not a string, return as-is so numbers/booleans compare normally.
  if (typeof value !== "string") return value;
  // Replace CRLF and CR with LF to unify line breaks across platforms/parsers.
  let out = value.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  // Replace non-breaking spaces with regular spaces so visually identical text matches.
  out = out.replace(/\u00A0/g, " ");
  // Trim to ignore inconsequential leading/trailing whitespace.
  return out.trim();
}

function isBlankCell(cell) {
  // cell: { v, f, t }
  const hasFormula = typeof cell.f === "string" && cell.f.startsWith("=");
  if (hasFormula) return false;
  if (cell.v === null || cell.v === "") return true;
  if (cell.t === "Empty") return true;
  return false;
}

function getCell(model, sidx, r, c) {
  const sh = model.sheets[sidx];
  if (!sh) return { v: null, f: null, t: "Empty" };
  if (r >= sh.rowCount || c >= sh.columnCount) return { v: null, f: null, t: "Empty" };
  const v = sh.values[r] && sh.values[r][c] !== undefined ? sh.values[r][c] : null;
  const f = (sh.formulas[r] && sh.formulas[r][c]) || null;
  const t = (sh.valueTypes[r] && sh.valueTypes[r][c]) || "Empty";
  return { v, f, t };
}

function classifyCell(a, b) {
  const aBlank = isBlankCell(a);
  const bBlank = isBlankCell(b);
  if (!aBlank && bBlank) return CODE_ADD;
  if (aBlank && !bBlank) return CODE_REMOVE;
  if (aBlank && bBlank) return CODE_NONE;

  // Both have something
  const af = normFormula(a.f);
  const bf = normFormula(b.f);
  if (af !== bf) {
    return CODE_FORMULA; // any formula text difference (including one side no formula)
  }
  // Same formula text; compare values
  // Normalize strings by trimming; numbers/booleans compare directly
  const av = normTextForCompare(a.v);
  const bv = normTextForCompare(b.v);
  if (av === bv) return CODE_NONE;
  // If both are literals (no formula), treat as explicit change (orange)
  if (af === "" /* and bf === "" by equality above */) return CODE_FORMULA;
  // Otherwise, same non-empty formula: calculated value-only change (yellow)
  return CODE_VALUE;
}

export function diffWorkbooks(curr, base) {
  const byNameCurr = new Map(curr.sheets.map((s, i) => [s.name, i]));
  const byNameBase = new Map(base.sheets.map((s, i) => [s.name, i]));

  const allNames = new Set([...curr.sheets.map((s) => s.name), ...base.sheets.map((s) => s.name)]);

  const bySheet = {};
  const sheetStatus = {};
  const summary = { total: { add: 0, remove: 0, value: 0, formula: 0, changedSheets: 0 } };

  for (const name of allNames) {
    const ai = byNameCurr.get(name);
    const bi = byNameBase.get(name);
    if (ai === undefined && bi !== undefined) {
      sheetStatus[name] = "removed";
      continue;
    }
    if (ai !== undefined && bi === undefined) {
      sheetStatus[name] = "added";
      continue;
    }
    // Both present
    const as = curr.sheets[ai];
    const bs = base.sheets[bi];
    // Determine absolute-space bounding box covering both used ranges
    const aRowOff = Math.max(0, as.rowOffset || 0);
    const aColOff = Math.max(0, as.colOffset || 0);
    const bRowOff = Math.max(0, bs.rowOffset || 0);
    const bColOff = Math.max(0, bs.colOffset || 0);
    const aRows = Math.max(0, as.rowCount || 0);
    const aCols = Math.max(0, as.columnCount || 0);
    const bRows = Math.max(0, bs.rowCount || 0);
    const bCols = Math.max(0, bs.columnCount || 0);
    const baseRow = Math.min(aRowOff, bRowOff);
    const baseCol = Math.min(aColOff, bColOff);
    const endRow = Math.max(aRowOff + aRows, bRowOff + bRows);
    const endCol = Math.max(aColOff + aCols, bColOff + bCols);
    const rows = Math.max(0, endRow - baseRow);
    const cols = Math.max(0, endCol - baseCol);
    const cells = new Uint8Array(rows * cols);
    let add = 0,
      remove = 0,
      value = 0,
      formula = 0;
    for (let rAbs = baseRow; rAbs < baseRow + rows; rAbs++) {
      for (let cAbs = baseCol; cAbs < baseCol + cols; cAbs++) {
        // Map absolute coordinate to local indices within each model's used range
        const ar = rAbs - aRowOff;
        const ac = cAbs - aColOff;
        const br = rAbs - bRowOff;
        const bc = cAbs - bColOff;
        const aCell = (ar >= 0 && ac >= 0 && ar < aRows && ac < aCols) ? getCell(curr, ai, ar, ac) : { v: null, f: null, t: "Empty" };
        const bCell = (br >= 0 && bc >= 0 && br < bRows && bc < bCols) ? getCell(base, bi, br, bc) : { v: null, f: null, t: "Empty" };
        const code = classifyCell(aCell, bCell);
        if (code !== CODE_NONE) {
          const rr = rAbs - baseRow;
          const cc = cAbs - baseCol;
          const idx = rr * cols + cc;
          cells[idx] = code;
          if (code === CODE_ADD) add++;
          else if (code === CODE_REMOVE) remove++;
          else if (code === CODE_VALUE) value++;
          else if (code === CODE_FORMULA) formula++;
        }
      }
    }
    const changed = add + remove + value + formula;
    sheetStatus[name] = changed > 0 ? "modified" : "unchanged";
    if (changed > 0) summary.total.changedSheets++;
    summary.total.add += add;
    summary.total.remove += remove;
    summary.total.value += value;
    summary.total.formula += formula;
    bySheet[name] = { rows, cols, rowBase: baseRow, colBase: baseCol, cells, counts: { add, remove, value, formula, changed } };
  }

  return { bySheet, sheetStatus, summary, codes: { CODE_NONE, CODE_ADD, CODE_REMOVE, CODE_VALUE, CODE_FORMULA } };
}
