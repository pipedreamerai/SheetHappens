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
