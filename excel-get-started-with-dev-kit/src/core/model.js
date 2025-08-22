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

    // Load used ranges (valuesOnly) for all sheets first.
    // valuesOnly=true ignores formatting-only regions (e.g., conditional formats),
    // preventing Mac Excel from shifting the used range start to A1.
    const usedRanges = items.map((ws) => ws.getUsedRangeOrNullObject(true));
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
      let rowOffset = 0; // zero-based start row of used range relative to A1
      let colOffset = 0; // zero-based start column of used range relative to A1

      if (rowCount && columnCount && !used.isNullObject) {
        const v = used.values || [];
        const f = used.formulas || [];
        const t = used.valueTypes || [];
        values = v.slice(0, rowCount).map((r) => (r || []).slice(0, columnCount));
        formulas = f
          .slice(0, rowCount)
          .map((r) => (r || []).slice(0, columnCount).map((c) => (typeof c === "string" ? c : null)));
        valueTypes = t.slice(0, rowCount).map((r) => (r || []).slice(0, columnCount));
        // Derive the top-left offset from the used range address (e.g., 'Sheet1!B2:D10')
        // We only need the starting cell to translate our local arrays back to absolute A1 coordinates.
        try {
          let addr = used.address;
          if (Array.isArray(addr)) addr = addr[0];
          if (typeof addr === "string" && addr) {
            const excl = addr.lastIndexOf("!"); // strip sheet name if present
            const local = excl >= 0 ? addr.slice(excl + 1) : addr;
            const first = local.includes(":") ? local.split(":")[0] : local;
            const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(first);
            if (m) {
              const letters = m[1].toUpperCase();
              const rowNum = parseInt(m[2], 10);
              let c = 0;
              for (let i = 0; i < letters.length; i++) {
                c = c * 26 + (letters.charCodeAt(i) - 64);
              }
              // Convert to zero-based
              colOffset = Math.max(0, c - 1);
              rowOffset = Math.max(0, rowNum - 1);
            }
          }
        } catch (_) {
          // If parsing fails, keep offsets at 0 which corresponds to A1
        }
      }

      model.sheets.push({
        name: ws.name,
        rowCount,
        columnCount,
        // Offsets allow us to map the 0-based indices of the used-range arrays
        // back to absolute worksheet coordinates (A1 origin), so diffs overlay correctly
        rowOffset,
        colOffset,
        values,
        formulas,
        valueTypes,
      });
    }

    return model;
  });
}
