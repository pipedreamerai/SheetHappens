/* eslint-disable office-addins/load-object-before-read */
/* global Excel */

// Build a lightweight, serializable snapshot of the current workbook.
// Options:
// - includeHidden: include hidden sheets (default: false)
// - maxCellsPerSheet: cap cells per sheet for safety/perf (default: null = unlimited)
export async function buildWorkbookModel(options = {}) {
  const { includeHidden = false, maxCellsPerSheet = null } = options;

  return Excel.run(async (context) => {
    const wb = context.workbook;
    const worksheets = wb.worksheets;
    worksheets.load("items/name,items/visibility");
    await context.sync();

    const items = (worksheets.items || []).filter(
      (ws) => includeHidden || ws.visibility !== Excel.SheetVisibility.hidden
    );

    const model = {
      name: "CurrentWorkbook",
      sheets: [],
    };

    // Load used ranges for all sheets first to avoid syncing inside the loop
    const usedRanges = items.map((ws) => ws.getUsedRangeOrNullObject());
    usedRanges.forEach((r) => r.load(["rowCount", "columnCount", "values", "formulas", "valueTypes", "address"]));
    await context.sync();

    for (let idx = 0; idx < items.length; idx++) {
      const ws = items[idx];
      const used = usedRanges[idx];

      let rowCount = used.isNullObject ? 0 : used.rowCount || 0;
      let columnCount = used.isNullObject ? 0 : used.columnCount || 0;

      // Safety cap: limit rows if over the max cell threshold
      if (rowCount && columnCount && maxCellsPerSheet && rowCount * columnCount > maxCellsPerSheet) {
        rowCount = Math.max(1, Math.floor(maxCellsPerSheet / Math.max(1, columnCount)));
      }

      let values = [];
      let formulas = [];
      let valueTypes = [];

      if (rowCount && columnCount && !used.isNullObject) {
        const v = used.values || [];
        const f = used.formulas || [];
        const t = used.valueTypes || [];
        values = v.slice(0, rowCount).map((r) => (r || []).slice(0, columnCount));
        formulas = f
          .slice(0, rowCount)
          .map((r) => (r || []).slice(0, columnCount).map((c) => (typeof c === "string" ? c : null)));
        valueTypes = t.slice(0, rowCount).map((r) => (r || []).slice(0, columnCount));
      }

      model.sheets.push({
        name: ws.name,
        rowCount,
        columnCount,
        values,
        formulas,
        valueTypes,
      });
    }

    return model;
  });
}
