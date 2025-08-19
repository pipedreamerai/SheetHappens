## Downloadable Archive (.xlsx): Plan (macOS, Web, Windows)

Goal: Add a one-click "Download archive (.xlsx)" that saves the current workbook as a strict .xlsx file across Mac, Web, and Windows. Prefer an exact copy via the platform API; fall back to rebuilding a .xlsx from our `WorkbookModel` when needed. Each step is small, testable, and reversible.

### High-level approach

- Primary (exact copy): Use `Office.context.document.getFileAsync(Office.FileType.Compressed)` to stream the workbook as a zipped OOXML package and save it as `.xlsx`.
- Fallback (model rebuild): If Compressed is unavailable or fails (host limit, throttling), rebuild a faithful-but-not-perfect `.xlsx` from our `WorkbookModel` using the `xlsx` library (values + formulas; limited formatting fidelity).
- UX: Add a button in the task pane, show progress, and name files like `<WorkbookName>_YYYYMMDD_HHMMSS.xlsx`.

### Minimal file changes

- New: `src/core/archive.js` — host-agnostic archive helpers
- Edit: `src/taskpane/taskpane.html` — add a button ("Download archive (.xlsx)")
- Edit: `src/taskpane/taskpane.js` — wire the button and progress messaging
- Docs: this plan and a short README section

---

## Step-by-step implementation

### 1) Capability detection + filename utility

Edits: `src/core/archive.js` (new)

- Export small helpers:
  - `isCompressedExportSupported(): boolean` — feature-detect `Office.context.document.getFileAsync` and `Office.FileType.Compressed`.
  - `buildArchiveFilename(workbookName: string): string` — timestamped pattern `<WorkbookName>_YYYYMMDD_HHMMSS.xlsx`.
  - `getWorkbookNameSafe(): Promise<string>` — read workbook name via Excel API; sanitize for file names.

Acceptance: Temporary dev log prints capability flags and a sample filename for a small workbook.
Rollback: Delete the new file and imports.

### 2) Primary exporter: Compressed → Blob(.xlsx)

Edits: `src/core/archive.js`

- Implement `exportCompressedAsBlob(progressCb?): Promise<{ blob: Blob, filename: string }>`:
  - Call `Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, cb)`.
  - On success, iterate `file.sliceCount` with `file.getSliceAsync(index, cb)`; accumulate `slice.data` (ArrayBuffer) into an array.
  - After all slices, `file.closeAsync()`; construct `new Blob(buffers, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })`.
  - Derive filename using `getWorkbookNameSafe()` + timestamp util. Call `progressCb` as slices progress.

Acceptance (manual):

- Click temporary dev button (or run via console) to export a small workbook; browser prompts to save `.xlsx`; file opens in Excel and matches the original sheets/formatting.
  Rollback: Guard behind a feature flag and skip calling it.

### 3) Fallback exporter: Rebuild from `WorkbookModel` using `xlsx`

Edits: `src/core/archive.js`

- Implement `exportFromModelAsBlob(model, progressCb?): Promise<{ blob: Blob, filename: string }>`:
  - Use `xlsx` to create a workbook in-memory.
  - For each visible sheet in `model.sheets`, write cell values and formulas (if `f` present, set the formula; else set the value).
  - Generate an ArrayBuffer via `XLSX.write(wb, { bookType: 'xlsx', type: 'array' })` and wrap in a Blob.
  - Filename from the same utility.

Notes:

- This preserves values/formulas but not rich formatting, objects, or macros.
- It ensures a strict `.xlsx` even if the primary path is unsupported.

Acceptance (manual):

- Force fallback (temporarily disable primary) and export; file opens; formulas recalc; sheet names/values match.
  Rollback: Remove the fallback call site; keep function for future use.

### 4) Orchestrator + UI wiring

Edits: `src/core/archive.js`, `src/taskpane/taskpane.js`, `src/taskpane/taskpane.html`

- Add `exportCurrentWorkbookXlsx({ preferCompressed: true }): Promise<{ ok: boolean, reason?: string }>` that:
  1. Tries `exportCompressedAsBlob`; on failure, logs and falls back to model rebuild using `buildWorkbookModel` + `exportFromModelAsBlob`.
  2. Triggers download using a safe cross-host method:
     - Preferred: createObjectURL + hidden `<a download>` click.
     - Fallback for legacy: `navigator.msSaveOrOpenBlob` if present.
- UI: Add a button "Download archive (.xlsx)" under Archive/Snapshots.
- Wire handler:
  - Disable button while running; show "Exporting… (n/N)" progress as slices load.
  - On success: "Archive downloaded: <filename>".
  - On failure: one-line error message, suggest retry or fallback reason.

Acceptance (manual):

- Click button on a small workbook; `.xlsx` downloads and opens; name pattern correct.
  Rollback: Hide the button; keep code behind a feature flag.

### 5) Cross-host verification matrix

No code, just manual checks per host:

- Excel for Web (Edge/Chrome): primary path should work; download via anchor should prompt Save. Verify identical workbook.
- Excel for Mac (latest): WebView download via anchor should prompt Save. Verify file opens and matches.
- Excel for Windows (latest): WebView2 download via anchor or `msSaveOrOpenBlob` should prompt Save/Open. Verify file opens and matches.

If primary path fails on a host: confirm fallback path (model rebuild) produces a valid `.xlsx` and warn about formatting loss.

### 6) Guardrails and limits

Edits: `src/core/archive.js` (logging + thresholds)

- Detect large files: if `sliceCount` is huge (e.g., > 20k slices) or estimated size exceeds a safe limit, show a friendly message and offer to try the fallback model rebuild.
- Always call `file.closeAsync()` even on errors.
- Handle user cancellation or throttling with clear, single-line messages.

Acceptance (manual):

- Simulate an error (disconnect dev tools or force a thrown error) and confirm we recover with a readable message.
  Rollback: Reduce strictness of thresholds.

### 7) Docs and small polish

- Add a README section: what the feature does, supported hosts, and limitations of the model-rebuild fallback.
- Add a feature flag in settings (optional): enable/disable fallback rebuild for strict environments.

Acceptance: Docs reflect behavior; toggle visible in code if implemented.
Rollback: Remove docs section and flag.

---

## Pseudocode snippets (for clarity only; implement in files noted above)

```javascript
// src/core/archive.js
export async function exportCompressedAsBlob(onProgress) {
  return new Promise((resolve, reject) => {
    try {
      const opt = { sliceSize: 65536 };
      Office.context.document.getFileAsync(Office.FileType.Compressed, opt, async (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(res.error && res.error.message || 'getFileAsync failed'));
        const file = res.value; // { size, sliceCount }
        const buffers = [];
        let fetched = 0;
        const getSlice = (i) => {
          file.getSliceAsync(i, (s) => {
            if (s.status !== Office.AsyncResultStatus.Succeeded) {
              try { file.closeAsync(() => {}); } catch (_) {}
              return reject(new Error(s.error && s.error.message || 'getSliceAsync failed'));
            }
            buffers.push(s.value.data); // ArrayBuffer
            fetched++;
            if (onProgress) onProgress({ fetched, total: file.sliceCount });
            if (fetched === file.sliceCount) {
              try { file.closeAsync(() => {}); } catch (_) {}
              const blob = new Blob(buffers, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
              const name = await getWorkbookNameSafe();
              resolve({ blob, filename: buildArchiveFilename(name) });
            } else {
              getSlice(fetched);
            }
          });
        };
        getSlice(0);
      });
    } catch (e) { reject(e); }
  });
}

export async function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
```

---

## Manual verification checklist

- Small workbook (few sheets): export via primary path; open in Excel; verify formatting, values, formulas, sheet names.
- Medium workbook: confirm progress updates and download completes; file opens.
- Force fallback: temporarily disable primary in code; export; verify values/formulas present; accept formatting loss.
- Error path: simulate failure; confirm single-line error and that UI re-enables the button.
- Hosts: test on Web, Mac, Windows; document any host-specific notes.

## Risks and mitigations

- Very large files may exceed practical limits for Compressed export. Mitigate with thresholds and the model-rebuild fallback.
- Desktop WebView download quirks. Mitigate with both anchor `download` and `msSaveOrOpenBlob` fallback when available.
- Rebuilt `.xlsx` lacks full fidelity. Make this explicit in UI messaging when fallback is used.

## Done criteria

- Button downloads a strict `.xlsx` on all three hosts.
- Progress displays; errors are concise.
- Fallback rebuild path works and is clearly communicated when used.
- Code isolated in `src/core/archive.js`, small UI edits only; clean and maintainable.
