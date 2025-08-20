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
  await logLinesToSheet([
    `Compare marker: after cacheDiffForLazyApply`,
  ], 'Compare');
  await logLinesToSheet([
    `Compare start: baseline='${baseName}', sheets=${Object.keys(diff.bySheet||{}).length}, backend=CF-only`
  ], 'Compare');
  diffEnabled = true;
  await logLinesToSheet([`Compare marker: before applyTabColors`], 'Compare');
  await applyTabColors(diff);
  // Immediately apply formatting for the currently active sheet
  await logLinesToSheet([`Compare marker: before applyDiffOnActivation`], 'Compare');
  // Run a smoke CF test on the current active sheet before lazy apply
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      await cfSmokeTest(context, ws);
    });
  } catch (_) { /* ignore */ }
  await applyDiffOnActivation();
      if (msg) msg.textContent = `Compared against ${baseName}: ${diff.summary.total.changedSheets} changed sheets`;
    } catch (e) {
      if (msg) msg.textContent = "Failed to compute diff: " + String(e && e.message ? e.message : e);
    }
  });
}

// Lightweight log sink in the task pane (appends to #dev-logs when present)
function appendToDevLogs(lines) {
  try {
    const el = document.getElementById('dev-logs');
    if (!el) return;
    const text = Array.isArray(lines) ? lines.join('\n') : String(lines);
    if (el.textContent && el.textContent.length) el.textContent += '\n' + text;
    else el.textContent = text;
    try { el.scrollTop = el.scrollHeight; } catch (_) { /* ignore */ }
  } catch (_) { /* ignore */ }
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
  // Console + taskpane output
  try { lines.forEach((l) => console.log(l)); } catch (_) {}
  appendToDevLogs(lines);
}

async function logLinesToSheet(lines, header = "Log") {
  const ts = new Date().toISOString();
  const payload = [
    `[${ts}] ${header}`,
    ...lines,
    "",
  ];
  // Console + taskpane output
  try { payload.forEach((l) => console.log(l)); } catch (_) {}
  appendToDevLogs(payload);
}

// In-context logger to avoid nested Excel.run; now writes to console
function appendLogsInContext(context, lines, header = "Log") {
  const ts = new Date().toISOString();
  // Normalize lines to strings and surface empty/falsy entries
  let normalized = Array.isArray(lines) ? lines.slice() : [lines];
  normalized = normalized.map((entry) => (entry == null ? "(empty)" : String(entry)));
  if (normalized.length === 0) normalized = ["(no details)"];
  const headerWithCount = `${header} (+${normalized.length} line${normalized.length === 1 ? '' : 's'})`;
  const payload = [
    `[${ts}] ${headerWithCount}`,
    ...normalized,
    "",
  ];
  // Console + taskpane output
  try { payload.forEach((l) => console.log(l)); } catch (_) {}
  appendToDevLogs(payload);
  // Preserve a Promise return for callers that await this function
  return Promise.resolve();
}

// ===== Lazy per-sheet diff formatting =====
const LAST_DIFF_KEY = 'cc_last_diff_cache_v1';
const APPLIED_ADDRESSES_KEY = 'cc_applied_addresses_v1';
const ORIGINAL_FILLS_KEY = 'cc_original_fills_v1'; // per-sheet map of original fills we overwrote
let lastDiffMem = null; // in-memory diff for immediate use
let diffEnabled = false; // whether to apply/generate overlays

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
  // Reset original fills tracking for a fresh run
  await saveSettingAsync(ORIGINAL_FILLS_KEY, {});
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
    await appendLogsInContext(context, [
      'Hooked worksheets.onActivated -> applyDiffOnActivation'
    ], 'Lazy Apply 1');
  }).catch(() => {});
}

async function applyDiffOnActivation() {
  try {
  if (!diffEnabled) return;
  const cached = restoreCachedDiff();
  if (!cached) return;
    await Excel.run(async (context) => {
      await appendLogsInContext(context, [
        `Lazy Apply: entering — diffEnabled=${String(diffEnabled)}, cached=${cached ? 'present' : 'missing'}`
      ], 'Lazy Apply 2');
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      active.load(['name']);
      await context.sync();
      const name = active.name;
      await appendLogsInContext(context, [
        `Lazy Apply: active='${name}'`
      ], 'Lazy Apply 3');
      // Run CF smoke test once per activation to validate host capabilities
      await cfSmokeTest(context, active);
      const s = cached.bySheet[name];
      if (!s) {
        await appendLogsInContext(context, [`Active sheet '${name}' has no diff entry`], "Lazy Apply 4");
        return;
      }
      // Build address runs per code and apply
  const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
  const already = applied[name];
      if (already && already.length) {
        await appendLogsInContext(context, [`Sheet '${name}' already applied; skipping`], "Lazy Apply 5");
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
        const deleted = await deleteTaggedOverlaysInRange(context, rect, new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR].map(normalizeColor)));
        await appendLogsInContext(context, [`Pre-clean CF deleted=${deleted}, usedRange=${rows}x${cols}`], "Lazy Apply 6");
      }
      const groups = buildAddressGroups(s);
      await appendLogsInContext(context, [
        `Entering CF apply for '${name}' with groups: add=${groups.add.length}, remove=${groups.remove.length}, value=${groups.value.length}, formula=${groups.formula.length}`
      ], "Lazy Apply 7");
      const sample = (arr) => arr.slice(0, 8).join(',');
      await appendLogsInContext(context, [
        `Applying groups for '${name}': add=${groups.add.length}, remove=${groups.remove.length}, value=${groups.value.length}, formula=${groups.formula.length}`,
        `Samples — add: ${sample(groups.add)}`,
        `Samples — remove: ${sample(groups.remove)}`,
        `Samples — value: ${sample(groups.value)}`,
        `Samples — formula: ${sample(groups.formula)}`,
      ], "Lazy Apply 20");
      await appendLogsInContext(context, [
        `Lazy Apply: calling applyGroupsToSheet — backend=CF-only`
      ], 'Lazy Apply 8');
      const appliedCounts = await applyGroupsToSheet(context, active, groups, async (msgs, hdr) => appendLogsInContext(context, msgs, hdr));
      applied[name] = [
        ...Object.values(groups).flat(),
      ];
      await saveSettingAsync(APPLIED_ADDRESSES_KEY, applied);
      await appendLogsInContext(context, [
        `Applied ${applied[name].length} ranges to '${name}'`,
        `Applied counts — add=${appliedCounts.add}, remove=${appliedCounts.remove}, value=${appliedCounts.value}, formula=${appliedCounts.formula}`,
      ], "Lazy Apply 9");
    });
  } catch (e) {
    await logLinesToSheet([`Error in lazy apply: ${String(e && e.message ? e.message : e)}`], "Lazy Apply Error");
  }
}

// Snapshot the original fills for a sheet's addresses we are about to overwrite.
// Stores only addresses with a non-null fill color and that haven't been captured yet.
/* eslint-disable office-addins/call-sync-before-read */
async function snapshotOriginalFillsForSheet(context, worksheet, sheetName, addresses) {
  try {
    if (!addresses || !addresses.length) return;
    const orig = getSetting(ORIGINAL_FILLS_KEY) || {};
    const existing = orig[sheetName] || [];
    const seen = new Set(existing.map((e) => e.addr));
    const ranges = [];
    const addrRefs = [];
    for (const addr of addresses) {
      if (seen.has(addr)) continue;
      // eslint-disable-next-line office-addins/call-sync-before-read
      const rg = worksheet.getRange(addr);
      rg.load(['format/fill/color']);
      ranges.push(rg);
      addrRefs.push(addr);
    }
    if (!ranges.length) return;
    // eslint-disable-next-line office-addins/no-context-sync-in-loop
    await context.sync();
    const toStore = [];
    for (let i = 0; i < ranges.length; i++) {
      const rg = ranges[i];
      try {
        const col = normalizeColor(rg.format.fill.color);
        if (col) toStore.push({ addr: addrRefs[i], color: col });
      } catch (_) { /* ignore */ }
    }
    if (toStore.length) {
      const updated = existing.concat(toStore);
      orig[sheetName] = updated;
      await saveSettingAsync(ORIGINAL_FILLS_KEY, orig);
    }
  } catch (_) {
    // best-effort; ignore errors
  }
}
/* eslint-enable office-addins/call-sync-before-read */

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
  try {
    if (logFn) await logFn([
      `CF backend: ENTER applyGroupsToSheet — add=${groups.add.length}, remove=${groups.remove.length}, value=${groups.value.length}, formula=${groups.formula.length}`
    ], 'CF Backend');
  } catch (_) {
    try { await logLinesToSheet([
      `CF backend (fallback): ENTER — add=${groups.add.length}, remove=${groups.remove.length}, value=${groups.value.length}, formula=${groups.formula.length}`
    ], 'CF Backend'); } catch (_) {}
  }
  // Apply conditional formats only; no direct fills
  const applyCF = async (addresses, color, label) => {
    try { await appendLogsInContext(context, [ `CF backend: category '${label}' addresses=${addresses ? addresses.length : 0}` ], 'CF Backend'); } catch (_) {
      try { await logLinesToSheet([ `CF backend (fallback): category '${label}' addresses=${addresses ? addresses.length : 0}` ], 'CF Backend'); } catch (_) {}
    }
    if (!addresses || addresses.length === 0) return 0;
    let created = 0;
    let sampled = 0;
    for (const addr of addresses) {
      try {
        const rg = worksheet.getRange(addr);
        let appliedType = "custom";
        try {
          if (logFn && sampled < 3) await logFn([`CF(${label}): try custom on ${addr}`], 'CF Backend');
          const cf = rg.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cf.custom.rule.formula = "=TRUE";
          // Prefer setting fill on the specific custom format; fall back as needed
          try { cf.custom.format.fill.setSolidColor(color); }
          catch (_) { try { cf.custom.format.fill.color = color; } catch (_) { try { cf.format.fill.color = color; } catch (_) { /* ignore */ } } }
        } catch (e1) {
          appliedType = "cellValue";
          try {
            if (logFn && sampled < 3) await logFn([`CF(${label}): custom failed on ${addr}: ${String(e1 && e1.message ? e1.message : e1)}`], 'CF Backend');
            if (logFn && sampled < 3) await logFn([`CF(${label}): try cellValue on ${addr}`], 'CF Backend');
            const cf2 = rg.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            try { cf2.cellValue.format.fill.setSolidColor(color); } catch (_) { try { cf2.cellValue.format.fill.color = color; } catch (_) { /* ignore */ } }
            cf2.cellValue.rule = { operator: Excel.ConditionalCellValueOperator.greaterThan, formula1: "-1" };
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
  const addN = await applyCF(groups.add, GREEN_COLOR, 'add');
  const remN = await applyCF(groups.remove, RED_COLOR, 'remove');
  const valN = await applyCF(groups.value, OVERLAY_COLOR, 'value');
  const frmN = await applyCF(groups.formula, ORANGE_COLOR, 'formula');
  await context.sync();
  try {
    if (logFn) await logFn([
      `CF backend: EXIT applyGroupsToSheet — created add=${addN}, remove=${remN}, value=${valN}, formula=${frmN}`
    ], 'CF Backend');
  } catch (_) {
    try { await logLinesToSheet([
      `CF backend (fallback): EXIT — created add=${addN}, remove=${remN}, value=${valN}, formula=${frmN}`
    ], 'CF Backend'); } catch (_) {}
  }
  return { add: addN, remove: remN, value: valN, formula: frmN };
}

function wireClearDiffFormatting() {
  const btn = document.getElementById('clear-diff-formatting');
  if (!btn) return;
  btn.addEventListener('click', async () => {
    const msg = document.getElementById('validation');
    try {
  // Stop diff generation and clear diff cache right away
  diffEnabled = false;
      await Excel.run(async (context) => {
        const wb = context.workbook;
        const wsCol = wb.worksheets;
        wsCol.load("items/name");
        await context.sync();
  const applied = getSetting(APPLIED_ADDRESSES_KEY) || {};
  const originalFills = getSetting(ORIGINAL_FILLS_KEY) || {};
  const colorSet = new Set([GREEN_COLOR, RED_COLOR, ORANGE_COLOR, OVERLAY_COLOR].map(normalizeColor));

        for (const ws of wsCol.items) {
          // Remove conditional formats with our colors only (no direct fills used)
          try {
            const addrs = Array.isArray(applied[ws.name]) ? applied[ws.name] : [];
            let totalDeleted = 0;
            if (addrs.length) {
              // Target known applied ranges first (covers CF-only cells not in used range)
              const sample = addrs.slice(0, Math.min(addrs.length, 200));
              for (const addr of sample) {
                try {
                  const r = ws.getRange(addr);
                  totalDeleted += await deleteTaggedOverlaysInRange(context, r, colorSet, { matchRuleTypes: true, brutal: true });
                } catch (_) { /* ignore */ }
              }
            }
            // Also sweep used range to catch any residuals (e.g., smoke test on A1)
            try {
              const u = ws.getUsedRange();
              u.load(['rowCount','columnCount']);
              // eslint-disable-next-line office-addins/no-context-sync-in-loop
              await context.sync();
              const rows = u.rowCount || 0;
              const cols = u.columnCount || 0;
              if (rows && cols) {
                const rect = ws.getRangeByIndexes(0,0,rows,cols);
                totalDeleted += await deleteTaggedOverlaysInRange(context, rect, colorSet, { matchRuleTypes: true });
              }
            } catch (_) { /* ignore */ }
            // Explicitly clear A1 (smoke test) with a narrow, brutal pass
            try {
              const a1 = ws.getRange('A1:A1');
              totalDeleted += await deleteTaggedOverlaysInRange(context, a1, colorSet, { matchRuleTypes: true, brutal: true });
            } catch (_) { /* ignore */ }
            try { await appendLogsInContext(context, [`${ws.name}: Deleted ${totalDeleted} conditional-format overlays (targeted+usedRange)`], "Stop Diff"); } catch (_) {}
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
        await saveSettingAsync(ORIGINAL_FILLS_KEY, {});
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
    try {
      const clearLogsBtn = document.getElementById('clear-dev-logs');
      if (clearLogsBtn) {
        clearLogsBtn.addEventListener('click', () => {
          try {
            const el = document.getElementById('dev-logs');
            if (el) el.textContent = '';
          } catch (_) { /* ignore */ }
        });
      }
    } catch (_) { /* ignore */ }
    // Startup banner: build/host/platform/version and requirement sets
    try {
      const stamp = String(Date.now());
      const host = (Office && Office.context && Office.context.host) || 'unknown-host';
      const platform = (Office && Office.context && Office.context.platform) || 'unknown-platform';
      const ver = (Office && Office.context && Office.context.diagnostics && Office.context.context && Office.context.diagnostics.version) || (Office && Office.context && Office.context.diagnostics && Office.context.diagnostics.version) || 'unknown-version';
      const rs12 = !!(Office && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported && Office.context.requirements.isSetSupported('ExcelApi','1.2'));
      const rs14 = !!(Office && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported && Office.context.requirements.isSetSupported('ExcelApi','1.4'));
      const rs17 = !!(Office && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported && Office.context.requirements.isSetSupported('ExcelApi','1.7'));
      const rs19 = !!(Office && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported && Office.context.requirements.isSetSupported('ExcelApi','1.9'));
      logLinesToSheet([
        `BUILD stamp=${stamp}`,
        `Host=${host}, Platform=${platform}, Version=${ver}`,
        `ExcelApi support: 1.2=${rs12}, 1.4=${rs14}, 1.7=${rs17}, 1.9=${rs19}`,
      ], 'Startup');
    } catch (_) { /* ignore banner log errors */ }
    wireArchiveSnapshot();
    populateSnapshotDropdown();
    wireUploadBaseline();
  wireRunCrossWorkbookSummary();
    initLazyFormatting();
    wireClearDiffFormatting();
  wireClearBaselines();
  } catch (e) {
    // ignore wiring errors
  }
});

// Simple CF smoke test on A1 to validate host supports CF add/format
async function cfSmokeTest(context, worksheet) {
  try {
    const rng = worksheet.getRange('A1:A1');
    const cfs = rng.conditionalFormats;
    const cf = cfs.add(Excel.ConditionalFormatType.custom);
    cf.custom.rule.formula = '=TRUE';
    try { cf.custom.format.fill.setSolidColor(OVERLAY_COLOR); }
    catch (_) { try { cf.custom.format.fill.color = OVERLAY_COLOR; } catch (_) { try { cf.format.fill.color = OVERLAY_COLOR; } catch (_) { /* ignore */ } } }
    cfs.load('items/type');
    await context.sync();
    const types = (cfs.items || []).map((it) => it.type).join(',');
    await appendLogsInContext(context, [
      `CF smoke: added custom CF to A1, count=${(cfs.items||[]).length}, types=${types}`
    ], 'CF Smoke');
  } catch (e) {
    await appendLogsInContext(context, [
      `CF smoke: failed on A1: ${String(e && e.message ? e.message : e)}`
    ], 'CF Smoke');
  }
}
