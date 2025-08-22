/* eslint-disable office-addins/load-object-before-read */
// Parse an uploaded .xlsx ArrayBuffer into a WorkbookModel compatible shape.
import * as XLSX from "xlsx";

function normalizeType(t) {
  switch (t) {
    case "n":
      return "Double";
    case "d":
      return "Double"; // treat dates as numbers for diff purposes
    case "s":
      return "String";
    case "b":
      return "Boolean";
    case "e":
      return "Error";
    default:
      return "Unknown";
  }
}

export function parseXlsxToModel(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);
  const wb = XLSX.read(data, { type: "array", cellDates: true, cellText: false });
  const sheetVis = (wb.Workbook && wb.Workbook.Sheets) || [];

  const model = {
    name: "UploadedWorkbook",
    sheets: [],
  };

  for (const name of wb.SheetNames) {
    // visibility: 0 visible, 1 hidden, 2 very hidden
    const visEntry = sheetVis.find((s) => s.name === name);
    const hidden = visEntry && typeof visEntry.Hidden === "number" ? visEntry.Hidden > 0 : false;
    if (hidden) continue; // skip hidden for MVP

    const ws = wb.Sheets[name];
    const ref = ws["!ref"];
    if (!ref) {
      model.sheets.push({ name, rowCount: 0, columnCount: 0, values: [], formulas: [], valueTypes: [] });
      continue;
    }
    const range = XLSX.utils.decode_range(ref);
    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;
    const rowOffset = Math.max(0, range.s.r); // zero-based starting row in worksheet coordinates
    const colOffset = Math.max(0, range.s.c); // zero-based starting column in worksheet coordinates

    const values = Array.from({ length: rows }, () => Array(cols).fill(null));
    const formulas = Array.from({ length: rows }, () => Array(cols).fill(null));
    const valueTypes = Array.from({ length: rows }, () => Array(cols).fill("Empty"));

    const addrRegex = /^[A-Z]+[0-9]+$/i;
    for (const key of Object.keys(ws)) {
      if (!addrRegex.test(key)) continue;
      const cell = ws[key];
      const addr = XLSX.utils.decode_cell(key);
      const r = addr.r - range.s.r;
      const c = addr.c - range.s.c;
      if (r < 0 || c < 0 || r >= rows || c >= cols) continue;
      const t = normalizeType(cell.t);
      valueTypes[r][c] = t;
      // Formula text (include leading '=') if present
      const f = cell.f ? `=${cell.f}` : null;
      formulas[r][c] = f;
      // Raw value; dates may be JS Date or number depending on cellDates
      values[r][c] = cell.v === undefined ? null : cell.v;
    }

    model.sheets.push({ name, rowCount: rows, columnCount: cols, rowOffset, colOffset, values, formulas, valueTypes });
  }

  return model;
}
