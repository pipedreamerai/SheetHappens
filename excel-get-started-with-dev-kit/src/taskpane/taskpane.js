/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideload = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideload) sideload.style.display = "none";
    if (appBody) appBody.classList.remove("is-hidden");
    // Initialize dropdowns and validation message.
    initSheetDropdowns();
  }
});

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
      if (runBtn) {
        const valid = Boolean(src.value) && Boolean(dst.value) && !same;
        runBtn.setAttribute("aria-disabled", String(!valid));
        runBtn.disabled = !valid; // stays disabled in this commit, but reflects validity
      }
    }

    src.addEventListener("change", validate);
    dst.addEventListener("change", validate);
    validate();
  }).catch((err) => {
    if (msg) msg.textContent = "Unable to enumerate worksheets: " + String(err && err.message ? err.message : err);
  });
}
