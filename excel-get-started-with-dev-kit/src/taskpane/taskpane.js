/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* eslint-disable prettier/prettier, office-addins/load-object-before-read */
/* global document, Office, Excel */

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
  }
});

const OVERLAY_TAG = 'CC_OVERLAY';
const OVERLAY_COLOR = '#FFF2CC'; // soft yellow as example overlay color

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
      const r1 = s1.getUsedRangeOrNullObject();
      const r2 = s2.getUsedRangeOrNullObject();
      r1.load(["rowCount", "columnCount"]);
      r2.load(["rowCount", "columnCount"]);
      await context.sync();

      // Normalize size: union bounds
      const rows = Math.max(r1.isNullObject ? 0 : r1.rowCount || 0, r2.isNullObject ? 0 : r2.rowCount || 0);
      const cols = Math.max(r1.isNullObject ? 0 : r1.columnCount || 0, r2.isNullObject ? 0 : r2.columnCount || 0);

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
        if (msg) msg.textContent = "Dry run is off. No formatting yet in this step.";
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
      const r1 = s1.getUsedRangeOrNullObject();
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
      const r1 = s1.getUsedRangeOrNullObject();
      // Load to allow access
      r1.load(["address"]);
      await context.sync();

      function removeOverlay(range) {
        if (!range || range.isNullObject) return;
        const cfs = range.conditionalFormats;
        cfs.load("items");
        return cfs;
      }

      const cf1 = removeOverlay(r1);
      await context.sync();

      function deleteTaggedOverlays(cfs) {
        if (!cfs || !cfs.items) return;
        // Queue loading of custom formulas for custom CFs
        const customs = [];
        cfs.items.forEach((cf) => {
          if (cf.type === Excel.ConditionalFormatType.custom) {
            cf.custom.rule.load("formula");
            customs.push(cf);
          }
        });
        return customs;
      }

      const customs1 = deleteTaggedOverlays(cf1);
      await context.sync();

      function purge(customs) {
        if (!customs) return;
        customs.forEach((cf) => {
          try {
            const formula = cf.custom.rule.formula || "";
            if (typeof formula === "string" && formula.indexOf(OVERLAY_TAG) !== -1) {
              cf.delete();
            }
          } catch (e) {
            // ignore and continue
          }
        });
      }

      purge(customs1);
      await context.sync();
      if (msg) msg.textContent = "Overlay removed on source sheet.";
    }).catch((err) => {
      if (msg) msg.textContent = "Failed to remove overlay: " + String(err && err.message ? err.message : err);
    });
  });
}
