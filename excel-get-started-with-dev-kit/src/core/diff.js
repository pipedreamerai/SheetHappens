// Diff two WorkbookModels (current vs baseline) and return per-sheet diffs and counts.

const CODE_NONE = 0;
const CODE_ADD = 1; // green
const CODE_REMOVE = 2; // red
const CODE_VALUE = 3; // yellow
const CODE_FORMULA = 4; // orange

function normFormula(f) {
  if (typeof f !== "string") return "";
  return f.trim().toUpperCase();
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
  const av = typeof a.v === "string" ? a.v.trim() : a.v;
  const bv = typeof b.v === "string" ? b.v.trim() : b.v;
  if (av !== bv) return CODE_VALUE;
  return CODE_NONE;
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
    const rows = Math.max(as.rowCount || 0, bs.rowCount || 0);
    const cols = Math.max(as.columnCount || 0, bs.columnCount || 0);
    const cells = new Uint8Array(rows * cols);
    let add = 0,
      remove = 0,
      value = 0,
      formula = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const a = getCell(curr, ai, r, c);
        const b = getCell(base, bi, r, c);
        const code = classifyCell(a, b);
        if (code !== CODE_NONE) {
          const idx = r * cols + c;
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
    bySheet[name] = { rows, cols, cells, counts: { add, remove, value, formula, changed } };
  }

  return { bySheet, sheetStatus, summary, codes: { CODE_NONE, CODE_ADD, CODE_REMOVE, CODE_VALUE, CODE_FORMULA } };
}
