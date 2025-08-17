/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* eslint-disable prettier/prettier, office-addins/load-object-before-read */
/* global document, Office, Excel, console, btoa, atob */
// eslint-disable-next-line no-unused-vars
import { buildWorkbookModel } from "../core/model";
import { saveSnapshot, listSnapshots, getSnapshot } from "../core/snapshot";
import { parseXlsxToModel } from "../core/import-xlsx";
import { diffWorkbooks } from "../core/diff";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideload = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideload) sideload.style.display = "none";
    if (appBody) appBody.classList.remove("is-hidden");
  // Initialize dropdowns and validation message.
  initSheetDropdowns();
  // Wire overlay actions.
  wireApplyOverlay();
  wireRemoveOverlay();
  wireRunCompareDryRun();
  wireDumpModel();
  wireArchiveSnapshot();
  populateSnapshotDropdown();
  wireInspectSnapshot();
  wireUploadBaseline();
  wireInspectUpload();
  wireRunCrossWorkbookSummary();
  wireResetTabColors();
  initLazyFormatting();
  wireClearDiffFormatting();
  }
});

const OVERLAY_TAG = 'CC_OVERLAY';
const OVERLAY_COLOR = '#FFF2CC'; // soft yellow as example overlay color
const GREEN_COLOR = '#C6EFCE';
const RED_COLOR = '#FFC7CE';
const ORANGE_COLOR = '#FFA500';

// Persist last overlay rect so we can precisely clear even if usedRange changes later.
const LAST_OVERLAY_KEY = 'cc_last_overlay_meta_v1';

function saveSettingAsync(key, value) {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(res.error || new Error('Failed to save settings.'));
      });
    } catch (e) {
      reject(e);
    }
  });
}

function getSetting(key) {
  try {
    return Office.context.document.settings.get(key);
  } catch (_) {
    return null;
  }
}

// Centralized helper to delete our overlays in a range, using only color-based matching.
// Runs two internal passes to handle cases where a first delete changes the items collection.
async function deleteTaggedOverlaysInRange(context, range, colorsSet) {
  if (!range) return 0;

  async function onePass() {
    const cfs = range.conditionalFormats;
    // Load types to filter custom formats
    cfs.load('items/type');
    await context.sync();
    const customItems = (cfs.items || []).filter((cf) => cf.type === Excel.ConditionalFormatType.custom);
    if (!customItems.length) return 0;
    // Load fill colors
    customItems.forEach((cf) => {
      try { cf.custom.format.fill.load('color'); } catch (_) { /* ignore */ }
    });
    await context.sync();
    const toDelete = customItems.filter((cf) => {
      try {
        const color = cf.custom && cf.custom.format && cf.custom.format.fill && cf.custom.format.fill.color;
        return typeof color === 'string' && colorsSet.has(color);
      } catch (_) { return false; }
    });
    toDelete.forEach((cf) => { try { cf.delete(); } catch (_) { /* ignore */ } });
    await context.sync();
    return toDelete.length;
  }

  let total = 0;
  try { total += await onePass(); } catch (e) {
    if (typeof console !== 'undefined' && console.warn) console.warn('deleteTaggedOverlaysInRange: pass1 failed', e);
  }
  // Second pass in case the collection changed after deletions (avoids double-click behavior)
  try { total += await onePass(); } catch (e) {
    if (typeof console !== 'undefined' && console.warn) console.warn('deleteTaggedOverlaysInRange: pass2 failed', e);
  }
  return total;
}

// Scan the active sheet for cells with direct fill colors matching our diff palette
// and clear them. Uses a conservative cap to avoid slow operations on large sheets.
async function clearDirectFillsByColor(context, ws, colorsSet, maxCells = 0) {
  const used = ws.getUsedRange(true); // include formatting-only cells
  used.load(["rowCount", "columnCount"]);
  await context.sync();
  const rows = used.rowCount || 0;
  const cols = used.columnCount || 0;
  const total = rows * cols;
  if (!rows || !cols) return { cleared: 0, skipped: false, total: 0 };
  if (maxCells && total > maxCells) return { cleared: 0, skipped: true, total };

  const addresses = [];
  // Keep batch size ~<= 5000 cells to limit request size
  const BATCH_ROWS = Math.max(1, Math.floor(5000 / Math.max(1, cols)));

  for (let r0 = 0; r0 < rows; r0 += BATCH_ROWS) {
    const r1 = Math.min(rows, r0 + BATCH_ROWS);
    const rowCells = [];
    for (let r = r0; r < r1; r++) {
      const arr = [];
      for (let c = 0; c < cols; c++) {
        const cell = ws.getCell(r, c);
        try { cell.format.fill.load("color"); } catch (_) { /* ignore */ }
        arr.push(cell);
      }
      rowCells.push(arr);
    }
  // We intentionally sync per batch to limit payload size; this is a controlled, bounded loop.
  // eslint-disable-next-line office-addins/no-context-sync-in-loop
  await context.sync();
    // Process loaded colors and group contiguous runs per row
    for (let i = 0; i < rowCells.length; i++) {
      const r = r0 + i;
      let c = 0;
      while (c < cols) {
        let color = null;
        try { color = rowCells[i][c].format.fill.color; } catch (_) { color = null; }
        if (!color || !colorsSet.has(color)) { c++; continue; }
        let c2 = c;
        while (c2 + 1 < cols) {
          let nextColor = null;
          try { nextColor = rowCells[i][c2 + 1].format.fill.color; } catch (_) { nextColor = null; }
          if (nextColor && colorsSet.has(nextColor)) c2++;
          else break;
        }
        addresses.push(toA1(r, c, r, c2));
        c = c2 + 1;
      }
    }
  }

  if (addresses.length) {
    try {
      const areas = ws.getRanges(addresses.join(","));
      try { areas.format.fill.clear(); } catch (_) {
        for (const a of addresses) { try { ws.getRange(a).format.fill.clear(); } catch (_) { /* ignore */ } }
      }
    } catch (_) {
      for (const a of addresses) { try { ws.getRange(a).format.fill.clear(); } catch (_) { /* ignore */ } }
    }
    await context.sync();
  }

  return { cleared: addresses.length, skipped: false, total };
}

function quoteSheetName(name) {
  const safe = String(name).replace(/'/g, "''");
  return `'${safe}'`;
}

function initSheetDropdowns() {
  const src = document.getElementById("source-sheet");
  const dst = document.getElementById("second-sheet");
  const msg = document.getElementById("validation");

  if (!(src && dst)) return;

  Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // Clear existing options except first placeholder
    function resetOptions(selectEl) {
      while (selectEl.options.length > 1) {
        selectEl.remove(1);
      }
    }
    resetOptions(src);
    resetOptions(dst);

    // Populate with worksheet names
    sheets.items.forEach((ws) => {
      const opt1 = document.createElement("option");
      opt1.value = ws.name;
      opt1.text = ws.name;
      src.appendChild(opt1);

      const opt2 = document.createElement("option");
      opt2.value = ws.name;
      opt2.text = ws.name;
      dst.appendChild(opt2);
    });

    // Preselect: source = active sheet, second = next sheet if available
    const active = sheets.getActiveWorksheet();
    active.load("name");
    await context.sync();
    src.value = active.name || "";
    if (sheets.items.length > 1) {
      const next = sheets.items.find((w) => w.name !== active.name);
      if (next) dst.value = next.name;
    }

    // Simple validation: cannot select the same sheet twice
    function validate() {
      const same = src.value && dst.value && src.value === dst.value;
      if (msg) {
        msg.textContent = same ? "Please choose two different sheets." : "";
      }
      const runBtn = document.getElementById("run-compare");
      const applyBtn = document.getElementById("apply-overlay");
      const removeBtn = document.getElementById("remove-overlay");
      const valid = Boolean(src.value) && Boolean(dst.value) && !same;
      const overlayValid = Boolean(src.value);
      if (runBtn) {
        runBtn.setAttribute("aria-disabled", String(!valid));
        runBtn.disabled = !valid; // stays disabled in this commit, but reflects validity
      }
      if (applyBtn) {
        applyBtn.setAttribute("aria-disabled", String(!overlayValid));
        applyBtn.disabled = !overlayValid;
      }
      if (removeBtn) {
        removeBtn.setAttribute("aria-disabled", String(!overlayValid));
        removeBtn.disabled = !overlayValid;
      }
    }

    src.addEventListener("change", validate);
    dst.addEventListener("change", validate);
    validate();
  }).catch((err) => {
    if (msg) msg.textContent = "Unable to enumerate worksheets: " + String(err && err.message ? err.message : err);
  });
}

function wireRunCompareDryRun() {
  const runBtn = document.getElementById("run-compare");
  if (!runBtn) return;
  runBtn.addEventListener("click", () => {
    const srcSel = document.getElementById("source-sheet");
    const dstSel = document.getElementById("second-sheet");
    const dry = document.getElementById("dry-run");
    const results = document.getElementById("dry-run-results");
    const msg = document.getElementById("validation");
    const sName = srcSel && srcSel.value ? srcSel.value : "";
    const dName = dstSel && dstSel.value ? dstSel.value : "";
    const doDryRun = dry && dry.checked;
    if (!(sName && dName) || sName === dName) {
      if (msg) msg.textContent = "Please select two different sheets.";
      return;
    }
    Excel.run(async (context) => {
      if (results) results.textContent = "";
      if (msg) msg.textContent = "";
      const wb = context.workbook;
      const s1 = wb.worksheets.getItem(sName);
      const s2 = wb.worksheets.getItem(dName);

      // Range mode: used range for now
  const u1 = s1.getUsedRange();
  const u2 = s2.getUsedRange();
  u1.load(['rowCount', 'columnCount']);
  u2.load(['rowCount', 'columnCount']);
  await context.sync();
  const u1Rows = u1.rowCount || 0;
  const u1Cols = u1.columnCount || 0;
  const u2Rows = u2.rowCount || 0;
  const u2Cols = u2.columnCount || 0;

      // Normalize size: union bounds
      const rows = Math.max(u1Rows, u2Rows);
      const cols = Math.max(u1Cols, u2Cols);

      function getRect(ws, rc, cc) {
        if (!rc || !cc) return null;
        return ws.getRangeByIndexes(0, 0, rc, cc);
      }

      const rect1 = getRect(s1, rows, cols);
      const rect2 = getRect(s2, rows, cols);
      if (rect1) rect1.load(["values", "formulas", "numberFormat", "text", "valueTypes"]);
      if (rect2) rect2.load(["values", "formulas", "numberFormat", "text", "valueTypes"]);
      await context.sync();

      if (!doDryRun) {
        // Presence/absence overlays on source sheet only
        if (rows && cols) {
          const rectSrc = s1.getRangeByIndexes(0, 0, rows, cols);
          // Load a lightweight property to ensure object is ready
          rectSrc.load(['address']);
          await context.sync();
          // Clean any existing overlay rules in the target rect (tag or known colors)
          const deleted = await deleteTaggedOverlaysInRange(context, rectSrc, new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR]));
          await appendLogsInContext(context, [
            `Presence/absence overlay pre-clean: deleted=${deleted} items in ${rectSrc.address}`
          ], 'Overlay');

          const other = quoteSheetName(dName);
          // Relative A1 reference to the top-left of rectSrc is A1
          const formulaGreen = `AND(NOT(ISBLANK(A1)), ISBLANK(${other}!A1), N("${OVERLAY_TAG}")=0)`;
          const formulaRed = `AND(ISBLANK(A1), NOT(ISBLANK(${other}!A1)), N("${OVERLAY_TAG}")=0)`;
          const cfG = rectSrc.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cfG.custom.rule.formula = formulaGreen;
          cfG.custom.format.fill.color = GREEN_COLOR;
          cfG.stopIfTrue = false;
          const cfR = rectSrc.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cfR.custom.rule.formula = formulaRed;
          cfR.custom.format.fill.color = RED_COLOR;
          cfR.stopIfTrue = false;

          // Commit 8 (revised): Orange for any differing values (including formulas), excluding the yellow case where formula text matches.
          // Rule: both cells non-blank and values differ, but not when both are formulas with identical FORMULATEXT.
          const formulaOrange = `AND(NOT(ISBLANK(A1)), NOT(ISBLANK(${other}!A1)), A1<>${other}!A1, NOT(AND(ISFORMULA(A1), ISFORMULA(${other}!A1), UPPER(TRIM(IFERROR(FORMULATEXT(A1),"")))=UPPER(TRIM(IFERROR(FORMULATEXT(${other}!A1),""))))), N("${OVERLAY_TAG}")=0)`;
          const cfO = rectSrc.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cfO.custom.rule.formula = formulaOrange;
          cfO.custom.format.fill.color = ORANGE_COLOR;
          cfO.stopIfTrue = false;

          // Commit 9: Yellow for same formula text but different results
          const formulaYellow = `AND(ISFORMULA(A1), ISFORMULA(${other}!A1), UPPER(TRIM(IFERROR(FORMULATEXT(A1),"")))=UPPER(TRIM(IFERROR(FORMULATEXT(${other}!A1),""))), A1<>${other}!A1, N("${OVERLAY_TAG}")=0)`;
          const cfY = rectSrc.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cfY.custom.rule.formula = formulaYellow;
          cfY.custom.format.fill.color = OVERLAY_COLOR; // light yellow
          cfY.stopIfTrue = false;

          await context.sync();
          await appendLogsInContext(context, [
            `Presence/absence overlays applied on ${sName} vs ${dName} in ${rectSrc.address}`
          ], 'Overlay');
          // Save last overlay rect for precise cleanup later
          try {
            await saveSettingAsync(LAST_OVERLAY_KEY, {
              sheet: sName,
              partner: dName,
              address: rectSrc.address,
              ts: Date.now(),
            });
          } catch (e) {
            if (typeof console !== 'undefined' && console.warn) {
              console.warn('Failed to persist last overlay rect', e);
            }
          }
          if (msg) msg.textContent = 'Presence/absence overlays applied to source sheet.';
        } else if (msg) {
          msg.textContent = 'Nothing to compare (empty ranges).';
        }
        return;
      }

      // Compute lightweight counts
      function countStats(r) {
        if (!r) return { cells: 0, blanks: 0, formulas: 0 };
        const vals = r.values || [];
        const forms = r.formulas || [];
        let cells = 0, blanks = 0, formulas = 0;
        for (let i = 0; i < vals.length; i++) {
          for (let j = 0; j < (vals[i] ? vals[i].length : 0); j++) {
            cells++;
            const v = vals[i][j];
            const f = forms[i] && forms[i][j];
            const hasFormula = typeof f === "string" && f.startsWith("=");
            if (hasFormula) formulas++;
            const isBlank = (!hasFormula) && (v === null || v === "");
            if (isBlank) blanks++;
          }
        }
        return { cells, blanks, formulas };
      }

      const s1Stats = rect1 ? countStats(rect1) : { cells: 0, blanks: 0, formulas: 0 };
      const s2Stats = rect2 ? countStats(rect2) : { cells: 0, blanks: 0, formulas: 0 };
      const summary = `Dry run — Rows x Cols: ${rows} x ${cols}. Source: ${s1Stats.cells} cells (${s1Stats.formulas} formulas, ${s1Stats.blanks} blanks). Second: ${s2Stats.cells} cells (${s2Stats.formulas} formulas, ${s2Stats.blanks} blanks).`;
      if (results) results.textContent = summary;
    }).catch((err) => {
      if (msg) msg.textContent = "Failed to run dry run: " + String(err && err.message ? err.message : err);
    });
  });
}

function wireApplyOverlay() {
  const btn = document.getElementById("apply-overlay");
  if (!btn) return;
  btn.addEventListener("click", () => {
    const srcSel = document.getElementById("source-sheet");
    const msg = document.getElementById("validation");
    const sName = srcSel && srcSel.value ? srcSel.value : "";
    if (!sName) {
      if (msg) msg.textContent = "Select a source sheet before applying overlay.";
      return;
    }
    Excel.run(async (context) => {
      const wb = context.workbook;
      const s1 = wb.worksheets.getItem(sName);
  const r1 = s1.getUsedRange();
  r1.load(["address"]);
      await context.sync();

      function addOverlay(range) {
        if (!range || range.isNullObject) return;
        const cfs = range.conditionalFormats;
        const cf = cfs.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = `OR(TRUE,N("${OVERLAY_TAG}"))`;
        cf.custom.format.fill.color = OVERLAY_COLOR;
        // Ensure it doesn't block other rules
        cf.stopIfTrue = false;
      }

      addOverlay(r1);
      await context.sync();
      // Save last overlay rect for precise cleanup later
      try {
        await r1.load('address');
        await context.sync();
        await saveSettingAsync(LAST_OVERLAY_KEY, {
          sheet: sName,
          partner: null,
          address: r1.address,
          ts: Date.now(),
        });
      } catch (e) {
        if (typeof console !== 'undefined' && console.warn) {
          console.warn('Failed to persist last overlay rect (overlay apply)', e);
        }
      }
      if (msg) msg.textContent = "Overlay applied to source sheet (used range).";
    }).catch((err) => {
      if (msg) msg.textContent = "Failed to apply overlay: " + String(err && err.message ? err.message : err);
    });
  });
}

function wireRemoveOverlay() {
  const btn = document.getElementById("remove-overlay");
  if (!btn) return;
  btn.addEventListener("click", () => {
    const srcSel = document.getElementById("source-sheet");
    const msg = document.getElementById("validation");
    const sName = srcSel && srcSel.value ? srcSel.value : "";
    if (!sName) {
      if (msg) msg.textContent = "Select a source sheet before removing overlay.";
      return;
    }
    Excel.run(async (context) => {
      const wb = context.workbook;
      const s1 = wb.worksheets.getItem(sName);

    let deletedTotal = 0;
  const colors = new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR]);

      // 1) Prefer the exact last overlay rect if available and still valid.
      const meta = getSetting(LAST_OVERLAY_KEY);
      if (meta && meta.sheet === sName && typeof meta.address === 'string') {
        try {
          const savedRange = s1.getRange(meta.address);
          savedRange.load(['address']);
          await context.sync();
          deletedTotal += await deleteTaggedOverlaysInRange(context, savedRange, colors);
        } catch (e) {
          if (typeof console !== 'undefined' && console.warn) {
            console.warn('Overlay remove: saved-range cleanup skipped', e);
          }
        }
      }

      // 2) Sweep the current source used range as a catch-all.
  const u1 = s1.getUsedRange();
  u1.load(['rowCount', 'columnCount']);
  await context.sync();
      if (u1.rowCount && u1.columnCount) {
        const rect = s1.getRangeByIndexes(0, 0, u1.rowCount, u1.columnCount);
        rect.load(['address']);
        await context.sync();
        deletedTotal += await deleteTaggedOverlaysInRange(context, rect, colors);
      }

      if (typeof console !== 'undefined' && console.log) {
        console.log('Overlay remove: deletedTotal =', deletedTotal);
      }
      await appendLogsInContext(context, [
        `Overlay remove on ${sName}: deletedTotal=${deletedTotal}`
      ], 'Overlay');
      if (msg) msg.textContent = deletedTotal > 0 ? 'Overlay removed on source sheet.' : 'No overlays found to remove.';
    }).catch((err) => {
      if (msg) msg.textContent = 'Failed to remove overlay: ' + String(err && err.message ? err.message : err);
    });
  });
}

function wireDumpModel() {
  const btn = document.getElementById("dump-model");
  if (!btn) return;
  btn.addEventListener("click", () => {
    const msg = document.getElementById("validation");
    if (msg) msg.textContent = "Building workbook model…";
    buildWorkbookModel({ includeHidden: false, maxCellsPerSheet: 200000 })
      .then(async (model) => {
        try {
          await dumpModelToLogsSheet(model);
          if (msg) msg.textContent = "Model written to 'logs' sheet.";
        } catch (e) {
          if (msg) msg.textContent = "Failed to write logs: " + String(e && e.message ? e.message : e);
        }
      })
      .catch((err) => {
        if (msg) msg.textContent = "Failed to build model: " + String(err && err.message ? err.message : err);
      });
  });
}

function wireArchiveSnapshot() {
  const btn = document.getElementById("archive-snapshot");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    const msg = document.getElementById("validation");
    if (msg) msg.textContent = "Creating snapshot…";
    try {
      const model = await buildWorkbookModel({ includeHidden: false, maxCellsPerSheet: 500000 });
      const name = `Snapshot ${new Date().toLocaleString()}`;
      const rec = await saveSnapshot(model, { name });
      await populateSnapshotDropdown();
      if (msg) msg.textContent = `Snapshot saved (${rec.sheetCount} sheets).`;
    } catch (e) {
      if (msg) msg.textContent = "Failed to save snapshot: " + String(e && e.message ? e.message : e);
    }
  });
}

async function populateSnapshotDropdown() {
  const sel = document.getElementById("baseline-snapshot");
  if (!sel) return;
  // Keep the first placeholder
  while (sel.options.length > 1) sel.remove(1);
  try {
    const items = await listSnapshots();
    items.forEach((it) => {
      const opt = document.createElement("option");
      const date = new Date(it.ts).toLocaleString();
      opt.value = it.id;
      opt.text = `${date} — ${it.name} (${it.sheetCount} sheets)`;
      sel.appendChild(opt);
    });
  } catch (e) {
    const msg = document.getElementById("validation");
    if (msg) msg.textContent = "Failed to load snapshots: " + String(e && e.message ? e.message : e);
  }
}

function wireInspectSnapshot() {
  const btn = document.getElementById("inspect-snapshot");
  const sel = document.getElementById("baseline-snapshot");
  if (!btn || !sel) return;
  btn.addEventListener("click", async () => {
    const msg = document.getElementById("validation");
    const id = sel.value;
    if (!id) {
      if (msg) msg.textContent = "Select a snapshot to inspect.";
      return;
    }
    if (msg) msg.textContent = "Loading snapshot…";
    try {
      const rec = await getSnapshot(id);
      if (!rec || !rec.model) {
        if (msg) msg.textContent = "Snapshot not found or has no model.";
        return;
      }
      await dumpModelToLogsSheet(rec.model);
      if (msg) msg.textContent = "Snapshot written to 'logs' sheet.";
    } catch (e) {
      if (msg) msg.textContent = "Failed to inspect snapshot: " + String(e && e.message ? e.message : e);
    }
  });
}

// In-memory stash for uploaded baselines this session
const uploadedBaselines = new Map(); // id -> { name, model }

function wireUploadBaseline() {
  const input = document.getElementById("upload-baseline");
  if (!input) return;
  input.addEventListener("change", async () => {
    const msg = document.getElementById("validation");
    const file = input.files && input.files[0];
    if (!file) return;
    if (msg) msg.textContent = "Parsing uploaded workbook…";
    try {
      const buf = await file.arrayBuffer();
      const model = parseXlsxToModel(buf);
      const id = `${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      uploadedBaselines.set(id, { name: file.name, model });
      addUploadedBaselineOption(id, file.name, model);
      if (msg) msg.textContent = `Uploaded baseline ready (${model.sheets.length} sheets).`;
    } catch (e) {
      if (msg) msg.textContent = "Failed to parse upload: " + String(e && e.message ? e.message : e);
    } finally {
      input.value = ""; // allow re-uploading same file
    }
  });
}

function addUploadedBaselineOption(id, name, model) {
  const sel = document.getElementById("baseline-uploaded");
  if (!sel) return;
  const opt = document.createElement("option");
  opt.value = id;
  opt.text = `${name} (${model.sheets.length} sheets)`;
  sel.appendChild(opt);
}

function wireInspectUpload() {
  const btn = document.getElementById("inspect-upload");
  const sel = document.getElementById("baseline-uploaded");
  if (!btn || !sel) return;
  btn.addEventListener("click", async () => {
    const msg = document.getElementById("validation");
    const id = sel.value;
    if (!id) {
      if (msg) msg.textContent = "No uploaded baseline selected.";
      return;
    }
    const entry = uploadedBaselines.get(id);
    if (!entry) {
      if (msg) msg.textContent = "Uploaded baseline not found in session.";
      return;
    }
    try {
      await dumpModelToLogsSheet(entry.model);
      if (msg) msg.textContent = "Uploaded baseline written to 'logs' sheet.";
    } catch (e) {
      if (msg) msg.textContent = "Failed to write uploaded baseline: " + String(e && e.message ? e.message : e);
    }
  });
}

function pickSelectedBaseline() {
  const snapSel = document.getElementById("baseline-snapshot");
  const upSel = document.getElementById("baseline-uploaded");
  const snapId = snapSel ? snapSel.value : "";
  const upId = upSel ? upSel.value : "";
  if (upId) {
    const entry = uploadedBaselines.get(upId);
    return entry ? { source: "upload", name: entry.name, model: entry.model } : null;
  }
  if (snapId) {
    return { source: "snapshot", id: snapId };
  }
  return null;
}

function wireRunCrossWorkbookSummary() {
  const btn = document.getElementById("run-xwb-summary");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    const msg = document.getElementById("validation");
    const choice = pickSelectedBaseline();
    if (!choice) {
      if (msg) msg.textContent = "Select a baseline (upload or snapshot) first.";
      return;
    }
    if (msg) msg.textContent = "Building models and computing diff…";
    try {
      const current = await buildWorkbookModel({ includeHidden: false, maxCellsPerSheet: 500000 });
      let baselineModel = null;
      let baseName = "Baseline";
      if (choice.source === "upload") {
        baselineModel = choice.model;
        baseName = choice.name;
      } else if (choice.source === "snapshot") {
        const rec = await getSnapshot(choice.id);
        if (!rec || !rec.model) throw new Error("Snapshot missing model");
        baselineModel = rec.model;
        baseName = rec.name || baseName;
      }
      const diff = diffWorkbooks(current, baselineModel);
  await writeSummaryToLogs(diff, baseName);
  // Cache diff for lazy per-sheet formatting
  await cacheDiffForLazyApply(diff);
  await applyTabColors(diff);
  // Immediately apply formatting for the currently active sheet
  await applyDiffOnActivation();
      if (msg) msg.textContent = `Compared against ${baseName}: ${diff.summary.total.changedSheets} changed sheets`;
    } catch (e) {
      if (msg) msg.textContent = "Failed to compute diff: " + String(e && e.message ? e.message : e);
    }
  });
}

async function writeSummaryToLogs(diff, baseName) {
  const ts = new Date().toISOString();
  const lines = [];
  lines.push(`[${ts}] Cross-workbook diff summary vs ${baseName}`);
  lines.push(
    `Total: +${diff.summary.total.add} / -${diff.summary.total.remove} / value ${diff.summary.total.value} / formula ${diff.summary.total.formula} | changed sheets: ${diff.summary.total.changedSheets}`
  );
  const names = Object.keys(diff.bySheet).sort();
  for (const n of names) {
    const s = diff.bySheet[n];
    if (!s || !s.counts) continue;
    const { add, remove, value, formula, changed } = s.counts;
    if (changed > 0) {
      lines.push(`- ${n}: +${add} / -${remove} / value ${value} / formula ${formula}`);
    } else {
      lines.push(`- ${n}: unchanged`);
    }
  }
  lines.push("", "----", "");

  await Excel.run(async (context) => {
    const wb = context.workbook;
    let logs = wb.worksheets.getItemOrNullObject("logs");
    logs.load(["name"]);
    await context.sync();
    if (logs.isNullObject) {
      logs = wb.worksheets.add("logs");
    }
  const used = logs.getUsedRange();
  used.load(["rowCount"]);
  await context.sync();
  const startRow = used.rowCount || 0;
    const rng = logs.getRangeByIndexes(startRow, 0, lines.length, 1);
    rng.values = lines.map((l) => [l]);
    try {
      const colA = logs.getRange("A:A");
      colA.format.columnWidth = 120;
    } catch (_) {
      // ignore formatting errors
    }
    await context.sync();
  });
}

async function logLinesToSheet(lines, header = "Log") {
  const ts = new Date().toISOString();
  const payload = [
    `[${ts}] ${header}`,
    ...lines,
    "",
  ];
  await Excel.run(async (context) => {
    const wb = context.workbook;
    let logs = wb.worksheets.getItemOrNullObject("logs");
    logs.load(["name"]);
    await context.sync();
    if (logs.isNullObject) logs = wb.worksheets.add("logs");
  const used = logs.getUsedRange();
  used.load(["rowCount"]);
  await context.sync();
  const startRow = used.rowCount || 0;
    const rng = logs.getRangeByIndexes(startRow, 0, payload.length, 1);
    rng.values = payload.map((l) => [l]);
    try { logs.getRange("A:A").format.columnWidth = 120; } catch (_) { /* ignore */ }
    await context.sync();
  });
}

// In-context logger to avoid nested Excel.run; writes lines to the 'logs' sheet using the provided context
function appendLogsInContext(context, lines, header = "Log") {
  const ts = new Date().toISOString();
  const payload = [
    `[${ts}] ${header}`,
    ...lines,
    "",
  ];
  const wb = context.workbook;
  let logs = wb.worksheets.getItemOrNullObject("logs");
  logs.load(["name"]);
  // We'll chain the operations after a sync in the caller for reliability
  return (async () => {
    await context.sync();
    if (logs.isNullObject) logs = wb.worksheets.add("logs");
  const used2 = logs.getUsedRange();
  used2.load(["rowCount"]);
  await context.sync();
  const startRow = used2.rowCount || 0;
    const rng = logs.getRangeByIndexes(startRow, 0, payload.length, 1);
    rng.values = payload.map((l) => [l]);
    try { logs.getRange("A:A").format.columnWidth = 120; } catch (_) { /* ignore */ }
    await context.sync();
  })();
}

// ===== Lazy per-sheet diff formatting =====
const LAST_DIFF_KEY = 'cc_last_diff_cache_v1';
const APPLIED_ADDRESSES_KEY = 'cc_applied_addresses_v1';
let lastDiffMem = null; // in-memory diff for immediate use

async function cacheDiffForLazyApply(diff) {
  // Store a compact version: bySheet with rows, cols, and a base64 of cells buffer
  const bySheet = {};
  for (const [name, s] of Object.entries(diff.bySheet)) {
    bySheet[name] = {
      rows: s.rows,
      cols: s.cols,
      cells: btoa(String.fromCharCode.apply(null, Array.from(s.cells))),
    };
  }
  // Keep in memory as well
  lastDiffMem = { bySheet: diff.bySheet };
  try {
    const approxBytes = Object.values(bySheet).reduce((sum, v) => sum + v.cells.length, 0);
    await saveSettingAsync(LAST_DIFF_KEY, { bySheet });
    await logLinesToSheet([
      `Cached diff to settings: sheets=${Object.keys(bySheet).length}, approxBytes=${approxBytes}`,
    ], "Diff Cache");
  } catch (e) {
    await logLinesToSheet([
      `Failed to cache diff to settings: ${String(e && e.message ? e.message : e)}`,
    ], "Diff Cache Error");
  }
  // Reset applied addresses tracking
  await saveSettingAsync(APPLIED_ADDRESSES_KEY, {});
}

function restoreCachedDiff() {
  // Prefer in-memory cache first
  if (lastDiffMem && lastDiffMem.bySheet) return lastDiffMem;
  const data = getSetting(LAST_DIFF_KEY);
  if (!data || !data.bySheet) return null;
  const bySheet = {};
  for (const [name, s] of Object.entries(data.bySheet)) {
    const bin = atob(s.cells);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    bySheet[name] = { rows: s.rows, cols: s.cols, cells: arr };
  }
  return { bySheet };
}

function initLazyFormatting() {
  // Attach a workbook sheet activation handler; apply formatting for the active sheet if cached diff exists
  Excel.run(async (context) => {
    const wb = context.workbook;
    wb.worksheets.onActivated.add(applyDiffOnActivation);
    await context.sync();
    await appendLogsInContext(context, [
      'Hooked worksheets.onActivated -> applyDiffOnActivation'
    ], 'Lazy Apply');
  }).catch(() => {});
}

async function applyDiffOnActivation() {
  try {
    const cached = restoreCachedDiff();
    if (!cached) {
      await logLinesToSheet(["No cached diff found; skipping apply"], "Lazy Apply");
      return;
    }
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      active.load(['name']);
      await context.sync();
      const name = active.name;
      const s = cached.bySheet[name];
      if (!s) {
        await appendLogsInContext(context, [`Active sheet '${name}' has no diff entry`], "Lazy Apply");
        return;
      }
      // Build address runs per code and apply
  const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
  const already = applied[name];
      if (already && already.length) {
        await appendLogsInContext(context, [`Sheet '${name}' already applied; skipping`], "Lazy Apply");
        return; // already applied this session
      }
      // Pre-clean any custom CF overlays that use our colors
  const u = active.getUsedRange();
  u.load(['rowCount','columnCount']);
  await context.sync();
  const rows = u.rowCount || 0;
  const cols = u.columnCount || 0;
      if (rows && cols) {
        const rect = active.getRangeByIndexes(0,0,rows,cols);
        const deleted = await deleteTaggedOverlaysInRange(context, rect, new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR]));
        await appendLogsInContext(context, [`Pre-clean CF deleted=${deleted}, usedRange=${rows}x${cols}`], "Lazy Apply");
      }
      const groups = buildAddressGroups(s);
      const sample = (arr) => arr.slice(0, 8).join(',');
      await appendLogsInContext(context, [
        `Applying groups for '${name}': add=${groups.add.length}, remove=${groups.remove.length}, value=${groups.value.length}, formula=${groups.formula.length}`,
        `Samples — add: ${sample(groups.add)}`,
        `Samples — remove: ${sample(groups.remove)}`,
        `Samples — value: ${sample(groups.value)}`,
        `Samples — formula: ${sample(groups.formula)}`,
      ], "Lazy Apply");
      const appliedCounts = await applyGroupsToSheet(context, active, groups, async (msgs, hdr) => appendLogsInContext(context, msgs, hdr));
      applied[name] = [
        ...Object.values(groups).flat(),
      ];
      await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied);
      await appendLogsInContext(context, [
        `Applied ${applied[name].length} ranges to '${name}'`,
        `Applied counts — add=${appliedCounts.add}, remove=${appliedCounts.remove}, value=${appliedCounts.value}, formula=${appliedCounts.formula}`,
      ], "Lazy Apply");
    });
  } catch (e) {
    await logLinesToSheet([`Error in lazy apply: ${String(e && e.message ? e.message : e)}`], "Lazy Apply Error");
  }
}

function buildAddressGroups(sheetDiff) {
  // Returns { add: [A1:A3,...], remove: [...], value: [...], formula: [...] }
  const { rows, cols, cells } = sheetDiff;
  const out = { add: [], remove: [], value: [], formula: [] };
  for (let r = 0; r < rows; r++) {
    let c = 0;
    while (c < cols) {
      const idx = r * cols + c;
      const code = cells[idx];
      if (!code) { c++; continue; }
      // extend horizontally for this code
      let c2 = c;
      while (c2 + 1 < cols && cells[r * cols + (c2 + 1)] === code) c2++;
      const addr = toA1(r, c, r, c2);
      if (code === 1) out.add.push(addr);
      else if (code === 2) out.remove.push(addr);
      else if (code === 3) out.value.push(addr);
      else if (code === 4) out.formula.push(addr);
      c = c2 + 1;
    }
  }
  return out;
}

function toA1(r1, c1, r2, c2) {
  // zero-based to A1; rows add 1
  function colName(c) {
    let x = c + 1;
    let s = '';
    while (x > 0) {
      const rem = (x - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      x = Math.floor((x - 1) / 26);
    }
    return s;
  }
  return `${colName(c1)}${r1 + 1}:${colName(c2)}${r2 + 1}`;
}

async function applyGroupsToSheet(context, worksheet, groups, logFn) {
  // Batch by category
  const apply = async (addresses, color) => {
    if (!addresses || addresses.length === 0) return;
    // Use comma-joined address string for getRanges
    const joined = addresses.join(',');
    try {
      const areas = worksheet.getRanges(joined);
      areas.format.fill.color = color;
    } catch (_) {
      // Fallback: per-address if getRanges(joined) not supported
      if (logFn) await logFn([`getRanges unsupported; falling back per-address (${addresses.length})`], "Lazy Apply");
      for (const addr of addresses) {
        try {
          const rg = worksheet.getRange(addr);
          rg.format.fill.color = color;
        } catch (_) { /* ignore */ }
      }
    }
  };
  await apply(groups.add, GREEN_COLOR);
  await apply(groups.remove, RED_COLOR);
  await apply(groups.value, OVERLAY_COLOR);
  await apply(groups.formula, ORANGE_COLOR);
  await context.sync();
  return {
    add: groups.add.length,
    remove: groups.remove.length,
    value: groups.value.length,
    formula: groups.formula.length,
  };
}

function wireClearDiffFormatting() {
  const btn = document.getElementById('clear-diff-formatting');
  if (!btn) return;
  btn.addEventListener('click', async () => {
    const msg = document.getElementById('validation');
    try {
      await Excel.run(async (context) => {
        const wb = context.workbook;
  const ws = wb.worksheets.getActiveWorksheet();
  ws.load(['name']);
  await context.sync();
  const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
  const name = ws.name;
        const addrs = (applied[name] || []);
        // If we tracked specific addresses, clear fills for those; fallback to color-based cleanup
        if (addrs.length) {
          try {
            // getRanges expects a comma-delimited A1 string, not an array
            const ranges = ws.getRanges(addrs.join(","));
            // RangeAreas.format is supported on recent requirement sets; clear if available
            try { ranges.format.fill.clear(); } catch (_) {
              // Fallback: per-address clear if RangeAreas.format not supported
              for (const a of addrs) {
                try { ws.getRange(a).format.fill.clear(); } catch (_) { /* ignore */ }
              }
            }
          } catch (_) {
            // Fallback: per-address clear if getRanges is not available/unsupported
            for (const a of addrs) {
              try { ws.getRange(a).format.fill.clear(); } catch (_) { /* ignore */ }
            }
          }
        }
        // Always run color-based conditional format cleanup as a safety net
  const u = ws.getUsedRange();
  u.load(['rowCount','columnCount']);
  await context.sync();
  const rows = u.rowCount || 0;
  const cols = u.columnCount || 0;
        if (rows && cols) {
          const rect = ws.getRangeByIndexes(0,0,rows,cols);
          await deleteTaggedOverlaysInRange(context, rect, new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR]));
          // Additionally, clear any direct fills that match our diff colors, even from previous runs
          try { await clearDirectFillsByColor(context, ws, new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR])); } catch (_) { /* ignore */ }
        }
        // Remove tracking for this sheet
        delete applied[name];
        await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied);
      });
      if (msg) msg.textContent = 'Cleared diff formatting for active sheet.';
    } catch (e) {
      if (msg) msg.textContent = 'Failed to clear diff formatting: ' + String(e && e.message ? e.message : e);
    }
  });
}

async function applyTabColors(diff) {
  // Priority: red (removed) > orange (formula) > yellow (value) > green (add) > default
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const wsCol = wb.worksheets;
    wsCol.load("items/name");
    await context.sync();
    for (const ws of wsCol.items) {
      const s = diff.bySheet[ws.name];
      if (!s || !s.counts) continue;
      const { add, remove, value, formula } = s.counts;
  let color = null;
  if (remove > 0) color = RED_COLOR;
  else if (formula > 0) color = ORANGE_COLOR;
  else if (value > 0) color = OVERLAY_COLOR; // yellow
  else if (add > 0) color = GREEN_COLOR;
      try {
        ws.tabColor = color;
      } catch (_) {
        // ignore if not supported
      }
    }
    await context.sync();
  });
}

function wireResetTabColors() {
  const btn = document.getElementById("reset-tab-colors");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    const msg = document.getElementById("validation");
    try {
      await Excel.run(async (context) => {
        const wsCol = context.workbook.worksheets;
        wsCol.load("items/name,items/tabColor");
        await context.sync();
        // First pass: null
        for (const ws of wsCol.items) {
          try { ws.tabColor = null; } catch (_) { /* ignore */ }
        }
        await context.sync();
        // Second pass: empty string as fallback for hosts that ignore null
        for (const ws of wsCol.items) {
          try { ws.tabColor = ""; } catch (_) { /* ignore */ }
        }
        await context.sync();
      });
      if (msg) msg.textContent = "Sheet tab colors reset.";
    } catch (e) {
      if (msg) msg.textContent = "Failed to reset tab colors: " + String(e && e.message ? e.message : e);
    }
  });
}

async function dumpModelToLogsSheet(model) {
  const ts = new Date().toISOString();
  // Prepare lines: header + JSON body split across lines to avoid cell length limits
  const header = [
    `[${ts}] WorkbookModel dump`,
    `Sheets: ${Array.isArray(model.sheets) ? model.sheets.length : 0}`,
    "",
  ];
  const json = JSON.stringify(model, null, 2);
  const bodyLines = json.split("\n");
  const lines = header.concat(bodyLines).concat(["", "----", ""]);

  await Excel.run(async (context) => {
    const wb = context.workbook;
    let logs = wb.worksheets.getItemOrNullObject("logs");
    logs.load(["name", "visibility"]);
    await context.sync();
    if (logs.isNullObject) {
      logs = wb.worksheets.add("logs");
      // Place at the end (optional)
      logs.position = wb.worksheets.count;
    }
  const used = logs.getUsedRange();
  used.load(["rowCount"]);
    await context.sync();
  const startRow = used.rowCount || 0;
    if (!lines.length) return;
    const rng = logs.getRangeByIndexes(startRow, 0, lines.length, 1);
    rng.values = lines.map((l) => [l]);
    // Make column A wide enough for readability (optional)
    try {
      const colA = logs.getRange("A:A");
      colA.format.columnWidth = 120;
    } catch (_) {
      // ignore formatting errors
    }
    await context.sync();
  });
}
