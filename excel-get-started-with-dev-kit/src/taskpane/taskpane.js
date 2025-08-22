/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* eslint-disable prettier/prettier, office-addins/load-object-before-read, office-addins/call-sync-before-read */
/* global document, Office, Excel, btoa, atob */
// eslint-disable-next-line no-unused-vars
import { buildWorkbookModel } from "../core/model";
import { saveSnapshot, listSnapshotsByWorkbook, getSnapshot, deleteSnapshot } from "../core/snapshot";
import { parseXlsxToModel } from "../core/import-xlsx";
import { diffWorkbooks } from "../core/diff";

// Diff colors and overlay tag used for identification/cleanup
const OVERLAY_COLOR = '#FFF2CC'; // yellow
const GREEN_COLOR = '#C6EFCE'; // added
const RED_COLOR = '#FFC7CE'; // removed
const ORANGE_COLOR = '#FFA500'; // formula change

// Persisted settings helpers
function saveSettingAsync(key, value) {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync((res) => {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res && res.error ? res.error.message : 'Failed to save settings'));
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

// Per-workbook identity (persists with the file, including Save As)
const WORKBOOK_ID_KEY = 'cc_workbook_id_v1';
function genGuid() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}
async function getOrCreateWorkbookId() {
  let id = getSetting(WORKBOOK_ID_KEY);
  if (!id) {
    id = genGuid();
    await saveSettingAsync(WORKBOOK_ID_KEY, id);
  }
  return id;
}

// Cleanup helpers
function normalizeColor(c) {
  return c == null ? null : String(c).trim().toUpperCase();
}

// Attempts to delete conditional formats within range that we applied.
// Deletion criteria:
// - If a fill color is readable: matches any in colors (Set of uppercase hex strings)
// - Optional rule-based match (when options.matchRuleTypes === true):
//   custom.rule.formula === '=TRUE' (ignoring case/leading '=') OR
//   cellValue.rule operator 'greaterThan' and formula1 '-1'
async function deleteTaggedOverlaysInRange(context, range, colors, options) {
  try {
    const cfs = range.conditionalFormats;
    try { cfs.load('items/type'); await context.sync(); } catch (_) { /* ignore */ }
    // Try to load format.fill.color when available
    for (const cf of cfs.items) {
      try { cf.load('format/fill/color,custom/format/fill/color,cellValue/format/fill/color'); } catch (_) { /* some CF types may not expose format */ }
      if (options && options.matchRuleTypes) {
        try { cf.load('custom/rule/formula,cellValue/rule'); } catch (_) { /* ignore */ }
      }
    }
    try { await context.sync(); } catch (_) { /* tolerate rule property load failures */ }
    let deleted = 0;
    for (const cf of cfs.items) {
      try {
        let shouldDelete = false;
        // Color-based tag match
        try {
          let col = null;
          try { col = cf.format && cf.format.fill ? cf.format.fill.color : null; } catch (_) { /* ignore */ }
          if (!col) {
            try { col = cf.custom && cf.custom.format && cf.custom.format.fill ? cf.custom.format.fill.color : null; } catch (_) { /* ignore */ }
          }
          if (!col) {
            try { col = cf.cellValue && cf.cellValue.format && cf.cellValue.format.fill ? cf.cellValue.format.fill.color : null; } catch (_) { /* ignore */ }
          }
          if (col && colors && colors.has(normalizeColor(col))) { shouldDelete = true; }
        } catch (_) { /* ignore */ }

        // Rule-based tag match (stricter; only when asked)
        if (!shouldDelete && options && options.matchRuleTypes) {
          try {
            if (cf.type === Excel.ConditionalFormatType.custom) {
              const f = (cf.custom && cf.custom.rule && cf.custom.rule.formula) || '';
              const norm = String(f).trim().replace(/^=/, '').toUpperCase();
              if (norm === 'TRUE') { shouldDelete = true; }
            } else if (cf.type === Excel.ConditionalFormatType.cellValue) {
              const rule = cf.cellValue && cf.cellValue.rule;
              const op = rule && rule.operator ? String(rule.operator).toLowerCase() : '';
              const f1 = rule && rule.formula1 ? String(rule.formula1) : '';
              if (op === 'greaterthan' && f1 === '-1') { shouldDelete = true; }
            }
          } catch (_) { /* ignore */ }
        }

        if (shouldDelete) { cf.delete(); deleted++; }
      } catch (_) { /* ignore */ }
    }
    await context.sync();
    // If nothing matched and caller allows, fall back to clearing all CF in this range
    if (deleted === 0 && options && options.brutal === true) {
      try {
        const coll = range.conditionalFormats;
        try { coll.load('items/type'); await context.sync(); } catch (_) {}
        try {
          coll.clearAll();
          await context.sync();
        } catch (_) { /* ignore */ }
        try { coll.load('items/type'); await context.sync(); } catch (_) {}
        // If still present or count unknown, try delete first item repeatedly as a last resort
        try {
          let loopDeletes = 0;
          for (let i = 0; i < 50; i++) {
            try {
              const cf0 = coll.getItemAt(0);
              cf0.delete();
              await context.sync();
              loopDeletes++;
            } catch (eIdx) {
              break;
            }
          }
          deleted = Math.max(deleted, loopDeletes);
        } catch (_) { /* ignore */ }
        // We don't know the exact count; report at least 1 to indicate action
        if (deleted === 0) deleted = 1;
      } catch (_) { /* ignore */ }
    }
    return deleted;
  } catch (_) {
    // On unexpected failure, attempt a last-chance brutal clear if allowed
    try {
      if (options && options.brutal === true) {
        const coll = range.conditionalFormats;
        try { coll.clearAll(); await context.sync(); } catch (_) {}
        return 1;
      }
    } catch (_) { /* ignore */ }
    return 0;
  }
}

// Batched deletion of tagged overlays across many addresses in a single worksheet.
// Loads CF collections and their properties in two syncs, then deletes in one batch.
async function deleteTaggedOverlaysForAddresses(context, worksheet, addresses, colors) {
  try {
    if (!addresses || !addresses.length) return 0;
    const entries = [];
    // First pass: queue loads of CF collections per address
    for (const addr of addresses) {
      try {
        const range = worksheet.getRange(addr);
        const cfs = range.conditionalFormats;
        try { cfs.load('items/type'); } catch (_) {}
        entries.push({ addr, range, cfs, items: [], deleted: 0 });
      } catch (_) { /* ignore bad addr */ }
    }
    try { await context.sync(); } catch (_) { /* tolerate */ }

    // Second pass: queue loads of properties for each CF item, track back to entry
    for (const entry of entries) {
      const list = (entry.cfs && entry.cfs.items) || [];
      entry.items = list;
      for (const cf of list) {
        try { cf.load('format/fill/color,custom/format/fill/color,cellValue/format/fill/color,custom/rule/formula,cellValue/rule'); } catch (_) {}
      }
    }
    try { await context.sync(); } catch (_) { /* tolerate */ }

    // Third pass: decide and delete, count per entry
    let totalDeleted = 0;
    for (const entry of entries) {
      for (const cf of entry.items) {
        try {
          let shouldDelete = false;
          // Color-based tag match
          try {
            let col = null;
            try { col = cf.format && cf.format.fill ? cf.format.fill.color : null; } catch (_) {}
            if (!col) { try { col = cf.custom && cf.custom.format && cf.custom.format.fill ? cf.custom.format.fill.color : null; } catch (_) {}
            }
            if (!col) { try { col = cf.cellValue && cf.cellValue.format && cf.cellValue.format.fill ? cf.cellValue.format.fill.color : null; } catch (_) {}
            }
            if (col && colors && colors.has(normalizeColor(col))) shouldDelete = true;
          } catch (_) {}

          // Rule-based tag match
          if (!shouldDelete) {
            try {
              if (cf.type === Excel.ConditionalFormatType.custom) {
                const f = (cf.custom && cf.custom.rule && cf.custom.rule.formula) || '';
                const norm = String(f).trim().replace(/^=/, '').toUpperCase();
                if (norm === 'TRUE') shouldDelete = true;
              } else if (cf.type === Excel.ConditionalFormatType.cellValue) {
                const rule = cf.cellValue && cf.cellValue.rule;
                const op = rule && rule.operator ? String(rule.operator).toLowerCase() : '';
                const f1 = rule && rule.formula1 ? String(rule.formula1) : '';
                if (op === 'greaterthan' && f1 === '-1') shouldDelete = true;
              }
            } catch (_) {}
          }

          if (shouldDelete) { cf.delete(); entry.deleted++; totalDeleted++; }
        } catch (_) { /* ignore */ }
      }
    }
    try { await context.sync(); } catch (_) {}

    // Final pass: for any entry with zero matched deletions, clear CF in that address as a targeted fallback
    for (const entry of entries) {
      try {
        if (entry.deleted === 0) {
          // Use our robust per-range deletion with rule match + brutal fallback
          try {
            await deleteTaggedOverlaysInRange(
              context,
              entry.range,
              colors,
              { matchRuleTypes: true, brutal: true }
            );
          } catch (_) {
            // Last-resort: try clearAll
            try { entry.range.conditionalFormats.clearAll(); } catch (_) {}
          }
        }
      } catch (_) { /* ignore */ }
    }
    try { await context.sync(); } catch (_) {}

    return totalDeleted;
  } catch (_) {
    return 0;
  }
}

// Removed aggressive color-based clearing of direct fills to avoid wiping user formatting.

function wireArchiveSnapshot() {
  const btn = document.getElementById("archive-snapshot");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    const msg = document.getElementById("validation");
    if (msg) msg.textContent = "Creating snapshot…";
    try {
  const workbookId = await getOrCreateWorkbookId();
      const model = await buildWorkbookModel({ includeHidden: false, maxCellsPerSheet: 500000 });
      const name = `Snapshot ${new Date().toLocaleString()}`;
  const rec = await saveSnapshot(model, { name, workbookId });
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
  const workbookId = await getOrCreateWorkbookId();
  const items = await listSnapshotsByWorkbook(workbookId);
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

// In-memory stash for uploaded baselines this session
const uploadedBaselines = new Map(); // id -> { name, model }

function wireUploadBaseline() {
  const input = document.getElementById("upload-baseline");
  if (!input) return;
  const chooseBtn = document.getElementById("choose-upload");
  if (chooseBtn) {
    chooseBtn.addEventListener("click", () => {
      try { input.click(); } catch (_) { /* ignore */ }
    });
  }
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

// Removed in simplified UI
// function wireInspectUpload() {}
function clearUploadedBaselinesUI() {
  // Clear the in-memory map and reset the uploaded dropdown to placeholder
  uploadedBaselines.clear();
  const sel = document.getElementById("baseline-uploaded");
  if (sel) {
    while (sel.options.length > 1) sel.remove(1);
    sel.selectedIndex = 0;
  }
}

function wireClearBaselines() {
  const btn = document.getElementById("clear-baselines");
  const container = document.getElementById("clear-baselines-confirm");
  if (!btn || !container) return;
  const reset = () => {
    container.classList.add("is-hidden");
    container.innerHTML = "";
  };
  btn.addEventListener("click", () => {
    // Render lightweight confirmation UI
    container.innerHTML = `
      <div class="confirm-panel" role="alertdialog" aria-labelledby="cbc-text">
        <div id="cbc-text" class="confirm-text">This will delete all your baselines (snapshots and uploads) for this workbook.</div>
        <div class="confirm-actions">
          <button id="cbc-confirm" class="ms-Button"><span class="ms-Button-label">Confirm</span></button>
          <button id="cbc-cancel" class="ms-Button"><span class="ms-Button-label">Cancel</span></button>
        </div>
      </div>`;
    container.classList.remove("is-hidden");
    const onCancel = () => reset();
    const onConfirm = async () => {
      const msg = document.getElementById("validation");
      if (msg) msg.textContent = "Clearing baselines…";
      try {
        const workbookId = await getOrCreateWorkbookId();
        const items = await listSnapshotsByWorkbook(workbookId);
        for (const it of items) {
          try { await deleteSnapshot(it.id); } catch (_) { /* ignore individual errors */ }
        }
        await populateSnapshotDropdown();
        clearUploadedBaselinesUI();
        if (msg) msg.textContent = "All baselines for this workbook deleted (snapshots and uploads).";
      } catch (e) {
        if (msg) msg.textContent = "Failed to clear baselines: " + String(e && e.message ? e.message : e);
      } finally {
        reset();
      }
    };
    const confirmBtn = document.getElementById("cbc-confirm");
    const cancelBtn = document.getElementById("cbc-cancel");
    if (confirmBtn) confirmBtn.addEventListener("click", onConfirm);
    if (cancelBtn) cancelBtn.addEventListener("click", onCancel);
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
    try { btn.classList.add('is-primary'); } catch (_) {}
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
  // Cache diff for lazy per-sheet formatting
  await cacheDiffForLazyApply(diff);
  
  diffEnabled = true;
  // Keep baseline model available for selection callouts
  lastBaselineModelMem = baselineModel;
  
  await applyTabColors(diff);
  // Apply full-sheet green overlays for sheets that are entirely new
  try { await applyAddedSheetOverlays(diff); } catch (_) {}
  // Immediately apply formatting for the currently active sheet
  
  await applyDiffOnActivation();
      if (msg) {
        try {
          const addedSheets = Object.entries(diff.sheetStatus || {}).filter(([, s]) => s === 'added').map(([n]) => n);
          const removedSheets = Object.entries(diff.sheetStatus || {}).filter(([, s]) => s === 'removed').map(([n]) => n);
          const parts = [`${diff.summary.total.changedSheets} changed sheets`];
          if (addedSheets.length) parts.push(`${addedSheets.length} added`);
          if (removedSheets.length) parts.push(`${removedSheets.length} removed`);
          msg.textContent = `Compared against ${baseName}: ${parts.join(' · ')}`;
        } catch (_) {
          msg.textContent = `Compared against ${baseName}: ${diff.summary.total.changedSheets} changed sheets`;
        }
      }
    } catch (e) {
      if (msg) msg.textContent = "Failed to compute diff: " + String(e && e.message ? e.message : e);
    }
  });
}

// Logging utilities removed for production UI

// ===== Lazy per-sheet diff formatting =====
const LAST_DIFF_KEY = 'cc_last_diff_cache_v1';
const APPLIED_ADDRESSES_KEY = 'cc_applied_addresses_v1';
const CREATED_REMOVED_SHEETS_KEY = 'cc_created_removed_sheets_v1';
let lastDiffMem = null; // in-memory diff for immediate use
let diffEnabled = false; // whether to apply/generate overlays
// Retain baseline model in memory for selection callouts
let lastBaselineModelMem = null;
// Track an active callout so we can clear it on selection changes
let activeCallout = { sheetName: null, address: null, weAddedValidation: false };
let selectionHandlerRef = null; // EventHandler removal token

async function cacheDiffForLazyApply(diff) {
  // Store a compact version: bySheet with rows, cols, and a base64 of cells buffer
  const bySheet = {};
  const createdRemoved = { added: [], removed: [] };
  for (const [name, s] of Object.entries(diff.bySheet)) {
    bySheet[name] = {
      rows: s.rows,
      cols: s.cols,
      rowBase: Math.max(0, s.rowBase || 0),
      colBase: Math.max(0, s.colBase || 0),
      cells: btoa(String.fromCharCode.apply(null, Array.from(s.cells))),
    };
  }
  // Track added/removed sheet names for messaging and overlays
  try {
    for (const [name, status] of Object.entries(diff.sheetStatus || {})) {
      if (status === 'added') createdRemoved.added.push(name);
      else if (status === 'removed') createdRemoved.removed.push(name);
    }
  } catch (_) {}
  // Keep in memory as well
  lastDiffMem = { bySheet: diff.bySheet };
  try {
    await saveSettingAsync(LAST_DIFF_KEY, { bySheet });
  } catch (_) { /* ignore */ }
  // Reset applied addresses tracking
  await saveSettingAsync(APPLIED_ADDRESSES_KEY, {});
  // Persist created/removed for later use
  try { await saveSettingAsync(CREATED_REMOVED_SHEETS_KEY, createdRemoved); } catch (_) {}
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
    bySheet[name] = { rows: s.rows, cols: s.cols, rowBase: Math.max(0, s.rowBase || 0), colBase: Math.max(0, s.colBase || 0), cells: arr };
  }
  return { bySheet };
}

async function clearCachedDiff() {
  try { await saveSettingAsync(LAST_DIFF_KEY, null); } catch (_) { /* ignore */ }
  try { await saveSettingAsync(APPLIED_ADDRESSES_KEY, {}); } catch (_) { /* ignore */ }
  lastDiffMem = null;
}

function initLazyFormatting() {
  // Attach a workbook sheet activation handler; apply formatting for the active sheet if cached diff exists
  Excel.run(async (context) => {
    const wb = context.workbook;
    wb.worksheets.onActivated.add(applyDiffOnActivation);
    await context.sync();
    
  }).catch(() => {});
}

async function applyDiffOnActivation() {
  try {
  if (!diffEnabled) return;
  const cached = restoreCachedDiff();
  if (!cached) return;
    await Excel.run(async (context) => {
      
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      active.load(['name']);
      await context.sync();
      const name = active.name;
      
      const s = cached.bySheet[name];
      if (!s) {
        // No per-cell diff for this sheet; it might be a newly added sheet (green tab) or unchanged.
        return;
      }
      // Build address runs per code and apply
  const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
  const already = applied[name];
      if (already && already.length) {
        
        return; // already applied this session
      }
      // Pre-clean overlays from our previous run:
      // 1) Prefer targeted deletion using stored rectangles for this sheet (fastest)
      // 2) Fallback to sweeping the used range once (slower, for first run or legacy leftovers)
  const prevRects = Array.isArray(applied[name]) ? applied[name] : [];
  const colorSet = new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR].map(normalizeColor));
  if (prevRects.length) {
        try { await deleteTaggedOverlaysForAddresses(context, active, prevRects, colorSet); } catch (_) {}
      } else {
        const u = active.getUsedRange(); // include formatting so CF-only overlays are sweepable
        u.load(['rowCount','columnCount']);
        await context.sync();
        const rows = u.rowCount || 0;
        const cols = u.columnCount || 0;
        if (rows && cols) {
          // Use the used range directly; do not assume it begins at A1.
          try { await deleteTaggedOverlaysInRange(context, u, colorSet, { matchRuleTypes: true }); } catch (_) {}
        }
      }
      const groups = buildAddressGroups(s);
      
      const sample = (arr) => arr.slice(0, 8).join(',');
      const appliedCounts = await applyGroupsToSheet(context, active, groups, null);
      applied[name] = [
        ...Object.values(groups).flat(),
      ];
      await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied);
      
    });
  } catch (e) {
    
  }
}

// Original direct-fill snapshotting has been removed; overlays are CF-only

function buildAddressGroups(sheetDiff) {
  // Returns merged rectangles per code as A1 ranges: { add: [A1:D5, ...], remove: [...], value: [...], formula: [...] }
  // Why rectangles? Creating one CF per row-run is expensive on large sheets. By merging identical
  // horizontal runs across consecutive rows, we create far fewer CF rules, making apply/clear much faster.
  const { rows, cols, cells } = sheetDiff;
  // Base offsets map our diff grid (which may start at B2, etc.) back to absolute A1 coordinates
  const rowBase = Math.max(0, sheetDiff.rowBase || 0); // newly-added: top-left absolute row
  const colBase = Math.max(0, sheetDiff.colBase || 0); // newly-added: top-left absolute col

  // Helper to collect horizontal segments for a given row and code
  function collectRowSegments(rowIndex, code) {
    // Build [c1, c2] segments where the code matches consecutively on this row
    const segs = [];
    let c = 0;
    while (c < cols) {
      const idx = rowIndex * cols + c;
      if (cells[idx] !== code) { c++; continue; }
      let c2 = c;
      while (c2 + 1 < cols && cells[rowIndex * cols + (c2 + 1)] === code) c2++;
      segs.push([c, c2]);
      c = c2 + 1;
    }
    return segs;
  }

  // Merge segments vertically into rectangles for a specific code
  function mergeRectanglesForCode(code) {
    const rectangles = []; // each: { r1, c1, r2, c2 }
    let prevRowOpen = new Map(); // key "c1,c2" -> rect
    for (let r = 0; r < rows; r++) {
      const segs = collectRowSegments(r, code);
      const nextRowOpen = new Map();
      // Try to extend previous rectangles when the same [c1,c2] segment appears in this row
      for (const [c1, c2] of segs) {
        const key = `${c1},${c2}`;
        if (prevRowOpen.has(key)) {
          // Extend existing rectangle downward
          const rect = prevRowOpen.get(key);
          rect.r2 = r;
          nextRowOpen.set(key, rect);
          prevRowOpen.delete(key);
        } else {
          // Start a new rectangle at this row
          nextRowOpen.set(key, { r1: r, c1, r2: r, c2 });
        }
      }
      // Any rectangles in prevRowOpen that did not continue should be finalized
      for (const rect of prevRowOpen.values()) rectangles.push(rect);
      prevRowOpen = nextRowOpen;
    }
    // Finalize remaining open rectangles from the last row
    for (const rect of prevRowOpen.values()) rectangles.push(rect);
    // Convert local rects to A1 using absolute offsets so overlays land in the right place
    return rectangles.map((rc) => toA1(rc.r1, rc.c1, rc.r2, rc.c2, rowBase, colBase));
  }

  // Build merged rectangles per code
  return {
    add: mergeRectanglesForCode(1),
    remove: mergeRectanglesForCode(2),
    value: mergeRectanglesForCode(3),
    formula: mergeRectanglesForCode(4),
  };
}

function toA1(r1, c1, r2, c2, rOff = 0, cOff = 0) {
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
  // Newly-added: apply absolute offsets so the A1 addresses point to real cells
  const rr1 = r1 + rOff;
  const cc1 = c1 + cOff;
  const rr2 = r2 + rOff;
  const cc2 = c2 + cOff;
  // Use absolute A1 refs to avoid any host re-anchoring quirks (e.g., Mac CF anchoring)
  return `$${colName(cc1)}$${rr1 + 1}:$${colName(cc2)}$${rr2 + 1}`;
}

async function applyGroupsToSheet(context, worksheet, groups, logFn) {
  
  // Apply conditional formats only; no direct fills
  const applyCF = async (addresses, color, label) => {
    
    if (!addresses || addresses.length === 0) return 0;
    let created = 0;
    let sampled = 0;
    for (const addr of addresses) {
      try {
        // Build range by absolute indexes to avoid any host quirks with A1 anchoring
        const rect = parseA1RangeToZeroBased(addr);
        const rg = rect
          ? worksheet.getRangeByIndexes(
              rect.r1,
              rect.c1,
              rect.r2 - rect.r1 + 1,
              rect.c2 - rect.c1 + 1
            )
          : worksheet.getRange(addr);
        let appliedType = "cellValue";
        // Try cellValue first to avoid Mac anchoring issues with custom rules
        try {
          if (logFn && sampled < 3) await logFn([`CF(${label}): try cellValue on ${addr}`], 'CF Backend');
          const cf2 = rg.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
          try { cf2.cellValue.format.fill.setSolidColor(color); } catch (_) { try { cf2.cellValue.format.fill.color = color; } catch (_) { /* ignore */ } }
          cf2.cellValue.rule = { operator: Excel.ConditionalCellValueOperator.greaterThan, formula1: "-1" };
        } catch (e1) {
          appliedType = "custom";
          try {
            if (logFn && sampled < 3) await logFn([`CF(${label}): cellValue failed on ${addr}: ${String(e1 && e1.message ? e1.message : e1)}`], 'CF Backend');
            if (logFn && sampled < 3) await logFn([`CF(${label}): try custom on ${addr}`], 'CF Backend');
            const cf = rg.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            cf.custom.rule.formula = "=TRUE";
            try { cf.custom.format.fill.setSolidColor(color); }
            catch (_) { try { cf.custom.format.fill.color = color; } catch (_) { try { cf.format.fill.color = color; } catch (_) { /* ignore */ } } }
          } catch (e2) {
            if (logFn) await logFn([`CF(${label}) failed on ${addr}: ${String(e2 && e2.message ? e2.message : e2)}`], 'CF Backend');
            continue;
          }
        }
        created++;
        if (logFn && sampled < 5) {
          await logFn([`CF(${label}) applied via ${appliedType} on ${addr}`], 'CF Backend');
          // Verify CF presence on this address (sample only)
          try {
            const cfs = rg.conditionalFormats;
            cfs.load("items/type");
            // eslint-disable-next-line office-addins/no-context-sync-in-loop
            await context.sync();
            const types = (cfs.items || []).map((it) => it.type).join(",");
            await logFn([`CF(${label}) verify on ${addr}: count=${(cfs.items || []).length}, types=${types}`], 'CF Backend');
          } catch (e3) {
            await logFn([`CF(${label}) verify failed on ${addr}: ${String(e3 && e3.message ? e3.message : e3)}`], 'CF Backend');
          }
          sampled++;
        }
      } catch (e) { /* ignore single-range errors */ }
    }
    if (logFn) await logFn([`Applied ${created} CF overlay(s) for ${label}`], 'CF Backend');
    return created;
  };
  // Prefer cellValue rule first on Mac: custom TRUE can occasionally anchor at A1 on some builds.
  // We enforce this by trying cellValue before custom in applyCF below.
  const addN = await applyCF(groups.add, GREEN_COLOR, 'add');
  const remN = await applyCF(groups.remove, RED_COLOR, 'remove');
  const valN = await applyCF(groups.value, OVERLAY_COLOR, 'value');
  const frmN = await applyCF(groups.formula, ORANGE_COLOR, 'formula');
  await context.sync();
  
  return { add: addN, remove: remN, value: valN, formula: frmN };
}

function wireClearDiffFormatting() {
  const btn = document.getElementById('clear-diff-formatting');
  if (!btn) return;
  btn.addEventListener('click', async () => {
    const msg = document.getElementById('validation');
    try { const startBtn = document.getElementById('run-xwb-summary'); if (startBtn) startBtn.classList.remove('is-primary'); } catch (_) {}
    try {
  // Stop diff generation and clear diff cache right away
  diffEnabled = false;
      await Excel.run(async (context) => {
        const wb = context.workbook;
        const wsCol = wb.worksheets;
        wsCol.load("items/name");
        await context.sync();
        // Clear any active selection callout
        try { await clearActiveCallout(); } catch (_) {}
  const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
  const colorSet = new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR].map(normalizeColor));

        for (const ws of wsCol.items) {
          // Remove conditional formats with our colors only (no direct fills used)
          try {
            const addrs = Array.isArray(applied[ws.name]) ? applied[ws.name] : [];
            let totalDeleted = 0;
            if (addrs.length) {
              // Target known rectangles in a single batched pass
              totalDeleted += await deleteTaggedOverlaysForAddresses(context, ws, addrs, colorSet);
            }
            // Always sweep used range once to catch any residual overlays (safe color/rule match only)
            try {
              const u = ws.getUsedRange(); // include formatting so CF-only overlays are sweepable
              u.load(['rowCount','columnCount']);
              await context.sync();
              const rows = u.rowCount || 0;
              const cols = u.columnCount || 0;
              if (rows && cols) {
                // Use the used range object directly; do not assume it begins at A1.
                // Enable brutal fallback to clear CF if tagged overlays aren't detectable post-revert.
                totalDeleted += await deleteTaggedOverlaysInRange(context, u, colorSet, { matchRuleTypes: true, brutal: true });
              }
            } catch (_) { /* ignore */ }
            
          } catch (_) { /* ignore */ }
          // Ensure worksheet gridlines are visible (CF does not hide them)
          try { ws.showGridlines = true; } catch (_) { /* ignore */ }
          // Clear tab color
          try { ws.tabColor = null; } catch (_) { /* ignore */ }
        }
        // Attempt a second pass to ensure tabs are cleared on hosts that ignore null
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();
        for (const ws of wsCol.items) { try { ws.tabColor = ""; } catch (_) { /* ignore */ } }
        // Wipe all tracking
        await saveSettingAsync(APPLIED_ADDRESSES_KEY, {});
      });
      // Now that we've used the stored addresses, clear the cached diff and in-memory copy
      await clearCachedDiff();
      if (msg) msg.textContent = 'Cleared diff formatting on all sheets.';
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
      const status = diff.sheetStatus ? diff.sheetStatus[ws.name] : undefined;
      let color = null;
      if (status === 'added') {
        // Entirely new sheet in current workbook
        color = GREEN_COLOR;
      } else if (status === 'removed') {
        // A sheet that existed in baseline but no longer exists now cannot be colored here;
        // we'll handle messaging and overlays separately.
      } else if (s && s.counts) {
        const { add, remove, value, formula } = s.counts;
        if (remove > 0) color = RED_COLOR;
        else if (formula > 0) color = ORANGE_COLOR;
        else if (value > 0) color = OVERLAY_COLOR; // yellow
        else if (add > 0) color = GREEN_COLOR;
      } else {
        // no diff info; leave default color
      }
      try {
        ws.tabColor = color;
      } catch (_) {
        // ignore if not supported
      }
    }
    await context.sync();
  });
}

// Apply a single green overlay covering the used range of any sheet marked as 'added'
async function applyAddedSheetOverlays(diff) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const wsCol = wb.worksheets;
    wsCol.load("items/name");
    await context.sync();
    const added = new Map(); // name -> worksheet
    for (const ws of wsCol.items) {
      const status = diff.sheetStatus ? diff.sheetStatus[ws.name] : undefined;
      if (status === 'added') added.set(ws.name, ws);
    }
    if (added.size === 0) return;
    // Load used ranges for all added sheets (valuesOnly to avoid formatting-only regions)
    const usedRanges = [];
    for (const [name, ws] of added) {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load(["address", "rowCount", "columnCount"]);
      usedRanges.push({ name, ws, u });
    }
    await context.sync();
    // Build and apply CF overlays, then record addresses
    const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
    for (const item of usedRanges) {
      try {
        const { name, ws, u } = item;
        if (u.isNullObject) continue;
        const rows = u.rowCount || 0;
        const cols = u.columnCount || 0;
        if (!rows || !cols) continue;
        // Compute absolute index rectangle from used range rows/cols (no A1 dependency)
        // Note: Used range is local to sheet; top-left index is (r1,c1) as parsed from address.
        let addr = u.address;
        if (Array.isArray(addr)) addr = addr[0];
        const rect = parseA1RangeToZeroBased(addr);
        if (!rect) continue;
        const r1 = rect.r1, c1 = rect.c1, h = rows, w = cols;
        // Apply CF overlay via index ranges and cellValue rule to avoid A1 anchoring
        try {
          const rg = ws.getRangeByIndexes(r1, c1, h, w);
          const cf2 = rg.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
          try { cf2.cellValue.format.fill.setSolidColor(GREEN_COLOR); } catch (_) { try { cf2.cellValue.format.fill.color = GREEN_COLOR; } catch (_) {} }
          cf2.cellValue.rule = { operator: Excel.ConditionalCellValueOperator.greaterThan, formula1: "-1" };
          // Track applied address for later cleanup (store A1 with absolutes for consistency)
          const localA1 = toA1(r1, c1, r1 + h - 1, c1 + w - 1);
          const arr = Array.isArray(applied[name]) ? applied[name] : [];
          if (!arr.includes(localA1)) arr.push(localA1);
          applied[name] = arr;
        } catch (_) { /* ignore single-sheet failure */ }
      } catch (_) { /* ignore */ }
    }
    try { await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied); } catch (_) {}
    await context.sync();
  });
}

// Removed in simplified UI
// function wireResetTabColors() {}

// Removed dumpModelToLogsSheet in simplified UI

// Wire up active controls when the task pane is ready
Office.onReady(() => {
  try {
  const sideload = document.getElementById('sideload-msg');
  const appBody = document.getElementById('app-body');
  if (sideload) sideload.classList.add('is-hidden');
  if (appBody) appBody.classList.remove('is-hidden');
    // Wire Clear Logs button
    // Clear Logs UI removed for production
    // Startup banner: build/host/platform/version and requirement sets
    // Startup banner logs removed for production
    wireArchiveSnapshot();
    populateSnapshotDropdown();
    wireUploadBaseline();
  wireRunCrossWorkbookSummary();
    initLazyFormatting();
    initSelectionCallouts();
    wireClearDiffFormatting();
  wireClearBaselines();
    // Revert selection button
    try {
      const rvBtn = document.getElementById('revert-selection');
      if (rvBtn) {
        
        rvBtn.addEventListener('click', async () => {
          try {
            
            await revertSelectedCellIfDiff();
          } catch (e) { }
        });
      } else { }
    } catch (e) { }
    // Hotkey removed per request (pane focus required). Use 'Revert Selection' button instead.
  } catch (e) {
    // ignore wiring errors
  }
});

// ===== Selection callout (New/Old) =====
function initSelectionCallouts() {
  try {
    Excel.run(async (context) => {
      const wb = context.workbook;
      async function wireActiveSheetSelection() {
        const active = wb.worksheets.getActiveWorksheet();
        active.load(["name"]);
        await context.sync();
        try {
          if (selectionHandlerRef && selectionHandlerRef.remove) {
            await selectionHandlerRef.remove();
          }
        } catch (_) { /* ignore remove errors */ }
        selectionHandlerRef = active.onSelectionChanged.add(async (event) => {
          try { await handleSelectionChanged(event); } catch (_) { /* ignore */ }
        });
        await context.sync();
      }
      // Wire now and on subsequent activations
      await wireActiveSheetSelection();
      wb.worksheets.onActivated.add(async () => { try { await handleActivationForSelection(); } catch (_) {} });
      await context.sync();
    }).catch(() => {});
  } catch (_) { /* ignore */ }
}

async function handleActivationForSelection() {
  try {
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      active.load(["name"]);
      await context.sync();
      try { if (selectionHandlerRef && selectionHandlerRef.remove) await selectionHandlerRef.remove(); } catch (_) {}
      selectionHandlerRef = active.onSelectionChanged.add(async (event) => {
        try { await handleSelectionChanged(event); } catch (_) { /* ignore */ }
      });
      await context.sync();
    });
  } catch (_) { /* ignore */ }
}

function parseA1ToZeroBased(addr) {
  // Accept forms like 'A1' only. Returns { row, col } zero-based or null
  if (!addr || typeof addr !== 'string') return null;
  let simple = addr.trim().toUpperCase();
  // Remove any sheet qualifier if present (e.g., 'Sheet1!A1')
  const excl = simple.lastIndexOf('!');
  let a = excl >= 0 ? simple.slice(excl + 1) : simple;
  // If a range like 'A1:A1' sneaks in, take the first cell
  if (a.indexOf(':') !== -1) a = a.split(':')[0];
  // Drop absolute markers like '$A$1'
  a = a.replace(/\$/g, '');
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  if (!m) return null;
  const colLetters = m[1];
  const rowNum = parseInt(m[2], 10);
  if (!rowNum || rowNum < 1) return null;
  let colNum = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { row: rowNum - 1, col: colNum - 1 };
}

function getBaselineCellValue(sheetName, r, c) {
  try {
    if (!lastBaselineModelMem) return { v: null, f: null, t: 'Empty' };
    const sh = (lastBaselineModelMem.sheets || []).find((s) => s && s.name === sheetName);
    if (!sh) return { v: null, f: null, t: 'Empty' };
    // Translate absolute worksheet coordinates to local used-range indices
    const rOff = Math.max(0, sh.rowOffset || 0);
    const cOff = Math.max(0, sh.colOffset || 0);
    const rl = r - rOff;
    const cl = c - cOff;
    if (rl < 0 || cl < 0 || rl >= (sh.rowCount || 0) || cl >= (sh.columnCount || 0)) return { v: null, f: null, t: 'Empty' };
    const v = sh.values && sh.values[rl] ? sh.values[rl][cl] : null;
    const f = sh.formulas && sh.formulas[rl] ? sh.formulas[rl][cl] : null;
    const t = sh.valueTypes && sh.valueTypes[rl] ? (sh.valueTypes[rl][cl] || 'Empty') : 'Empty';
    return { v, f, t };
  } catch (_) {
    return { v: null, f: null, t: 'Empty' };
  }
}

function formatValueForDisplay(cell) {
  if (!cell) return '';
  const f = typeof cell.f === 'string' && cell.f ? cell.f : null;
  if (f && f.startsWith('=')) return f;
  if (cell.v == null) return '';
  return String(cell.v);
}

function parseA1RangeToZeroBased(rangeA1) {
  try {
    if (!rangeA1) return null;
    let txt = String(rangeA1).trim();
    const excl = txt.lastIndexOf('!');
    if (excl >= 0) txt = txt.slice(excl + 1);
    const parts = txt.split(':');
    const norm = parts.length === 1 ? [parts[0], parts[0]] : [parts[0], parts[1]];
    const p = (cell) => {
      const c = cell.replace(/\$/g, '').toUpperCase();
      const m = /^([A-Z]+)(\d+)$/.exec(c);
      if (!m) return null;
      const colLetters = m[1];
      const rowNum = parseInt(m[2], 10);
      let col = 0; for (let i = 0; i < colLetters.length; i++) col = col * 26 + (colLetters.charCodeAt(i) - 64);
      return { row: rowNum - 1, col: col - 1 };
    };
    const a = p(norm[0]);
    const b = p(norm[1]);
    if (!a || !b) return null;
    return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
  } catch (_) { return null; }
}

function a1AddressContainsCell(addr, row, col) {
  try {
    const r = parseA1RangeToZeroBased(addr);
    if (!r) return false;
    return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
  } catch (_) { return false; }
}

function encodeUint8ToBase64(u8) {
  try { return btoa(String.fromCharCode.apply(null, Array.from(u8))); } catch (_) { return ''; }
}

async function clearActiveCallout() {
  try {
    if (!activeCallout || !activeCallout.sheetName || !activeCallout.address) return;
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(activeCallout.sheetName);
      const rg = ws.getRange(activeCallout.address);
      try {
        if (activeCallout.weAddedValidation && rg.dataValidation && rg.dataValidation.clear) {
          rg.dataValidation.clear();
        }
      } catch (_) { /* ignore */ }
      await context.sync();
    });
  } catch (_) { /* ignore */ }
  activeCallout = { sheetName: null, address: null, weAddedValidation: false };
}

async function handleSelectionChanged(event) {
  try {
    if (!diffEnabled) return;
    const cached = restoreCachedDiff();
    if (!cached) return;
    // Resolve active sheet and selection address
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const ws = wb.worksheets.getActiveWorksheet();
      ws.load(["name", "id"]);
      await context.sync();
      const sheetName = ws.name;
      // Clear previous callout if any
      try { await clearActiveCallout(); } catch (_) {}
      const addr = event && event.address ? event.address : null;
      const pos = parseA1ToZeroBased(addr);
      if (!pos) return; // only single-cell selections supported
      const s = cached.bySheet && cached.bySheet[sheetName];
      if (!s) return;
      const rows = s.rows || 0;
      const cols = s.cols || 0;
      const rBase = Math.max(0, s.rowBase || 0);
      const cBase = Math.max(0, s.colBase || 0);
      if (pos.row < rBase || pos.col < cBase || pos.row >= rBase + rows || pos.col >= cBase + cols) return;
      const rLocal = pos.row - rBase;
      const cLocal = pos.col - cBase;
      const code = s.cells[rLocal * cols + cLocal];
      if (!code) return; // unchanged
      // Read current cell value/formula
      const target = ws.getRange(addr);
      target.load(["values", "formulas"]);
      await context.sync();
      const currVal = (target.values && target.values[0] ? target.values[0][0] : null);
      const currF = (target.formulas && target.formulas[0] ? target.formulas[0][0] : null);
      const currCell = { v: currVal, f: typeof currF === 'string' ? currF : null };
      const baseCell = getBaselineCellValue(sheetName, pos.row, pos.col);
      // If this is a yellow cell (value-only change) but the value now equals baseline, remove the overlay
      try {
        if (code === 3) {
          const cv = (typeof currCell.v === 'string') ? currCell.v.trim() : currCell.v;
          const bv = (typeof baseCell.v === 'string') ? baseCell.v.trim() : baseCell.v;
          if (cv === bv) {
            // Targeted cleanup like revert path
            try {
              const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
              const sheetApplied = Array.isArray(applied[sheetName]) ? applied[sheetName] : [];
              const keep = [];
              let removedCount = 0;
              for (const a of sheetApplied) {
                if (a1AddressContainsCell(a, pos.row, pos.col)) {
                  try {
                    const rr = ws.getRange(a);
                    const deleted = await deleteTaggedOverlaysInRange(context, rr, new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR].map(normalizeColor)), { matchRuleTypes: true, brutal: true });
                    removedCount += deleted;
                  } catch (_) { /* ignore */ }
                } else {
                  keep.push(a);
                }
              }
              applied[sheetName] = keep;
              await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied);
              // Last-chance: clear CF only on this single cell
              try { const cellCF = target.conditionalFormats; cellCF.clearAll(); await context.sync(); } catch (_) { /* ignore */ }
              
            } catch (_) { /* ignore */ }
            return; // do not show tooltip
          }
        }
      } catch (_) { /* ignore */ }
      let newText = '';
      let oldText = '';
      if (code === 4) { // formula change
        newText = currCell.f && currCell.f.startsWith('=') ? currCell.f : (currCell.v == null ? '' : String(currCell.v));
        oldText = baseCell && typeof baseCell.f === 'string' && baseCell.f ? baseCell.f : (baseCell && baseCell.v != null ? String(baseCell.v) : '');
      } else if (code === 3) { // value change (same formula)
        newText = currCell.v == null ? '' : String(currCell.v);
        oldText = baseCell && baseCell.v != null ? String(baseCell.v) : '';
      } else if (code === 1) { // added in current
        newText = formatValueForDisplay(currCell);
        oldText = '';
      } else if (code === 2) { // removed from current
        newText = '';
        oldText = formatValueForDisplay(baseCell);
      }
      // If both strings are empty, do not show
      if (!newText && !oldText) {
        
        return;
      }
      // Respect existing data validation if present
      let alreadyHasValidation = false;
      try {
        const dv = target.dataValidation;
        dv.load(["rule/type", "prompt/showPrompt"]);
        await context.sync();
        alreadyHasValidation = !!(dv && dv.rule && dv.rule.type);
      } catch (_) { alreadyHasValidation = false; }
      if (alreadyHasValidation) return;
      try {
        const dv = target.dataValidation;
        try {
          dv.prompt = { showPrompt: true, title: 'New / Old', message: `New: ${newText}\nOld: ${oldText}` };
        } catch (_) {
          try { dv.inputMessage = { showInputMessage: true, title: 'New / Old', message: `New: ${newText}\nOld: ${oldText}` }; } catch (_) { /* ignore */ }
        }
        activeCallout = { sheetName, address: addr, weAddedValidation: true };
      } catch (e) { }
      await context.sync();
    });
  } catch (_) { /* ignore */ }
}

// Revert the currently selected area (single cell or multi-cell) to the baseline for green(1)/red(2)/orange(4) only
async function revertSelectedCellIfDiff() {
  try {
    if (!diffEnabled) {
      return;
    }
    const cached = restoreCachedDiff();
    if (!cached) { return; }
    if (!lastBaselineModelMem) { return; }
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const ws = wb.worksheets.getActiveWorksheet();
      ws.load(["name"]);
      const sel = wb.getSelectedRange();
      // Load selection details (address + shape). We won't rely on formulas/values here.
      sel.load(["address", "rowCount", "columnCount"]);
      await context.sync();
      const sheetName = ws.name;
      let addr = sel.address;
      if (Array.isArray(addr)) addr = addr[0];
      // Parse rectangular selection like 'A1:D5' to zero-based bounds
      const rect = parseA1RangeToZeroBased(addr);
      if (!rect) { return; }
      const sheetDiff = cached.bySheet && cached.bySheet[sheetName];
      if (!sheetDiff) { return; }
      const { rows, cols, cells } = sheetDiff;
      const rBase = Math.max(0, sheetDiff.rowBase || 0);
      const cBase = Math.max(0, sheetDiff.colBase || 0);
      // Clamp selection to the absolute bounds of the diff grid
      const absR1 = Math.max(rBase, rect.r1);
      const absC1 = Math.max(cBase, rect.c1);
      const absR2 = Math.min(rBase + rows - 1, rect.r2);
      const absC2 = Math.min(cBase + cols - 1, rect.c2);
      if (absR1 > absR2 || absC1 > absC2) { return; }

      // Queue cell edits in one batch: for each cell in selection, apply baseline
      // Only revert add/remove/formula-changed cells (codes 1,2,4); skip value-only (3)
      const changedCells = [];
      for (let rAbs = absR1; rAbs <= absR2; rAbs++) {
        for (let cAbs = absC1; cAbs <= absC2; cAbs++) {
          const rLocal = rAbs - rBase;
          const cLocal = cAbs - cBase;
          const code = cells[rLocal * cols + cLocal];
          if (!(code === 1 || code === 2 || code === 4)) continue;
          const base = getBaselineCellValue(sheetName, rAbs, cAbs);
          const baselineFormula = (typeof base.f === 'string' && base.f) ? base.f : null;
          const baselineValue = (base.v == null ? null : base.v);
          const cellRange = ws.getRangeByIndexes(rAbs, cAbs, 1, 1);
          // Prefer formulas when present; else set literal or blank
          if (baselineFormula) {
            try { cellRange.formulas = [[baselineFormula]]; } catch (_) {}
          } else if (baselineValue !== null) {
            try { cellRange.values = [[baselineValue]]; } catch (_) {}
          } else {
            try { cellRange.values = [[""]]; } catch (_) {}
          }
          changedCells.push({ r: rAbs, c: cAbs });
        }
      }

      // If nothing to do, return early
      if (!changedCells.length) { return; }

      // Force recalc once after the batch
      try { ws.calculate(Excel.CalculationType.recalculate); } catch (_) {}

      // Remove overlays only for changed cells (batched)
      try {
        const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
        const sheetApplied = Array.isArray(applied[sheetName]) ? applied[sheetName] : [];
        const toDeleteAddrs = new Set();
        const keep = [];
        for (const a of sheetApplied) {
          let intersects = false;
          for (const cc of changedCells) {
            if (a1AddressContainsCell(a, cc.r, cc.c)) { intersects = true; break; }
          }
          if (intersects) {
            toDeleteAddrs.add(a);
          } else {
            keep.push(a);
          }
        }
        // Execute deletions on intersecting CF ranges in one batch
        await deleteTaggedOverlaysForAddresses(
          context,
          ws,
          Array.from(toDeleteAddrs.values()),
          new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR].map(normalizeColor))
        );
        applied[sheetName] = keep;
        await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied);
      } catch (e) { }

      await context.sync();
    });
  } catch (e) {
    
  }
}

