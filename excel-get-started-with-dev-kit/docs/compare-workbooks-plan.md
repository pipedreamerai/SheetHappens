# Excel Workbook Versioning: Critique + Revised Plan (macOS, local-first)

This rewrites the existing plan with a leaner architecture, clearer contracts, and a commit-by-commit path that’s small, testable, and reversible. It preserves your current same-workbook compare and adds local archiving, cross-workbook diffing, lazy formatting, and sheet-tab colors. Cloud is deliberately last.

## Quick critique of the previous plan
- Good: local-first, keeps existing same-workbook CF rules, defers cloud, calls out lazy-load and tab colors.
- Gaps tightened here:
  - Contracts were not explicit enough for testing; we add clear input/output types and result shapes.
  - Diff application strategy needs batching with RangeAreas and rectangle grouping to avoid slow per-cell writes.
  - Archiving should rely on Office File APIs when available, with a graceful fallback to snapshot JSON; both are included.
  - Event wiring must include onAdded/onDeleted/onActivated to keep lazy-load robust as sheets change.
  - Explicit success criteria, rollbacks, and perf guardrails were light; these are now embedded per commit.

## MVP goals (macOS, no SharePoint/OneDrive)
- Archive current workbook locally via user action: JSON snapshot in IndexedDB and an optional .xlsx/.ooxml download.
- Compare current workbook vs a selected baseline (uploaded .xlsx or stored snapshot).
- Visualize differences using green/red/yellow/orange fills in the current workbook.
- Lazy-load formatting: only apply fills when a sheet becomes active.
- Sheet tab colors to quickly spot changes across sheets.

## Color semantics (align with current UI)
- Green `#C6EFCE`: Added (present in current, blank in baseline).
- Red `#FFC7CE`: Removed (blank in current, present in baseline).
- Orange `#FFA500`: Formula changed (FORMULATEXT differs; value may or may not differ).
- Yellow `#FFF2CC`: Same formula text, different evaluated result.
- Unchanged: no fill.

## Core contracts (tiny “APIs” for testability)
- WorkbookModel
  - name: string
  - sheets: SheetModel[]
- SheetModel
  - name: string
  - rowCount: number
  - columnCount: number
  - values: any[][]
  - formulas: (string|null)[][]
  - valueTypes: ("Empty"|"String"|"Double"|"Boolean"|"Error"|"Unknown")[][]
- Diff
  - bySheet: Record<string, CellDiffMap>
  - sheetStatus: Record<string, "added"|"removed"|"modified"|"unchanged">
- CellDiffMap (dense grid sized to max rows/cols of the pair)
  - rows: number, cols: number
  - cells: Uint8Array or number[][] with enum codes:
    - 0=none, 1=green(add), 2=red(remove), 3=yellow(value change), 4=orange(formula change)

Error modes: invalid workbook file, oversized sheets (guardrails), throttling, permission/user-cancel. Success: non-throwing functions return models/diffs; UI shows per-sheet counts.

## Architecture and file layout
- Keep `taskpane.js` for wiring and UI; extract logic into modules under `src/core/`:
  - `src/core/model.js` — build WorkbookModel from Excel runtime.
  - `src/core/import-xlsx.js` — parse uploaded .xlsx to WorkbookModel (SheetJS/xlsx).
  - `src/core/diff.js` — pure diff engine producing Diff + counts.
  - `src/core/format.js` — apply/clear fills using RangeAreas; build rectangles from CellDiffMap.
  - `src/core/snapshot.js` — IndexedDB CRUD for snapshots.
  - `src/core/events.js` — subscribe/unsubscribe to onActivated/onAdded/onDeleted.
- Shared constants (colors, keys) in `src/core/constants.js`.

## UX tweaks (task pane)
- Controls: “Archive snapshot”, “Choose baseline …” (upload or snapshot), “Run compare”, “Clear formatting”, toggle “Lazy apply by sheet”.
- Summary: total changes + per-sheet counts; click a sheet name to activate it.
- Non-blocking progress messages; fail gracefully with a single-line error in the pane.

## Commit-by-commit plan (each has acceptance + rollback)

0) baseline (existing): same-workbook CF compare
- Keep current behavior and overlay cleaning. Add a tiny constant module for shared color hexes.
- Acceptance: Existing green/red/orange/yellow CF still works.

1) feat(model): extract current workbook → WorkbookModel (done)
- Implement `model.js`: read visible sheets, usedRange, values, formulas, valueTypes; cap at N rows/cols via a setting (default unlimited, warn over 50k cells per sheet).
- Add a dev-only “Dump model” button to log JSON.
- Acceptance: Small workbook logs correct sizes and a sample of values/formulas.
- Rollback: delete `model.js` import and button.

2) feat(snapshot): IndexedDB snapshots (JSON) (done)
- `snapshot.js`: save/load/list/delete snapshots with id, name, ts, label.
- UI: “Archive snapshot” saves current model; list snapshots in baseline picker.
- Acceptance: Snapshot appears after click; persists across reloads.
- Rollback: hide button and ignore IndexedDB if unavailable.

3) feat(import): upload .xlsx baseline (SheetJS)
- Add `xlsx` dependency; file input → ArrayBuffer → WorkbookModel.
- Normalize used ranges (infer max rows/cols from non-empty cells).
- Acceptance: Upload test file; see sheets in picker with row/col counts.
- Rollback: hide upload; feature flag in settings.

4) feat(diff): pure diff engine
- `diff.js`: compare two WorkbookModels by sheet name.
- Cell rules: compute add/remove and value/formula-change using valueTypes, trimmed strings, and exact formula-text compare (case/whitespace-insensitive using TRIM+UPPER equivalent in JS).
- Output counts and sheetStatus (added/removed/modified/unchanged).
- Acceptance: Summary shows expected counts on crafted fixtures; add 2-3 unit tests for engine (Jest or lightweight harness).
- Rollback: keep engine behind flag; fall back to existing same-workbook CF.

5) feat(format): efficient apply/clear with RangeAreas
- Convert CellDiffMap to rectangles via row run-length encode → merge contiguous cells; build A1 addresses per category.
- Use `worksheet.getRanges(addressList)` or `RangeAreas` to batch set `format.fill.color` per category.
- Track applied addresses in document settings per sheet for reliable “Clear formatting”. Reuse color-based cleanup as safety net.
- Acceptance: On one sheet, apply/clear runs quickly (<1s on medium sheets) and leaves no orphan formatting.
- Rollback: disable format module; no changes committed to workbook.

6) feat(lazy): apply on activation only
- `events.js`: subscribe to `workbook.worksheets.onActivated`, also handle `onAdded/onDeleted` to keep wiring consistent.
- When a sheet activates and has a diff, apply fills if not already applied; do not pre-apply hidden/inactive sheets.
- Add toggle in UI; default ON.
- Acceptance: Switching sheets applies formatting on demand; Clear removes current sheet’s fills; toggling OFF applies all visible sheets.
- Rollback: unsubscribe events; keep manual Apply.

7) feat(tabs): sheet tab colors
- Compute a per-sheet most-severe category and set `worksheet.tabColor` accordingly.
- Priority: red > orange > yellow > green > default.
- Add a “Reset tab colors” action.
- Acceptance: Tabs change on compare and reset on command.
- Rollback: remove color set and reset.

8) feat(archive): downloadable copy of current workbook
- Try Office File APIs: export OOXML or compressed workbook; create a Blob and trigger user download named `Archive/<WorkbookName>_YYYYMMDD_HHMMSS.xlsx` (or .xml if OOXML only is available). Store the last chosen “Archive” folder name (not path) in settings.
- Fallback (if API not supported): snapshot JSON plus an instruction link to “Save As” in Excel UI.
- Acceptance: Button prompts a download; file opens in Excel.
- Rollback: disable button; keep snapshots only.

9) perf: guardrails and batching
- Chunk very large address lists, throttle `context.sync`, and early-abort with a helpful message if estimated writes exceed a safe limit.
- Setting: max cells per sheet for compare.
- Acceptance: Large-but-realistic books complete without throttling; message appears if limits exceeded.
- Rollback: increase limits or skip formatting when over threshold.

10) docs/tests: polish
- Add README sections for archive, upload, compare, lazy mode, and tab colors; document limitations.
- Expand diff unit tests (value vs formula, blanks, errors).
- Acceptance: Lint/tests pass; docs up to date.

11) cloud (optional last): OneDrive/SharePoint
- Add SSO; list and write to an Archive folder near the workbook; baseline picker can browse cloud archive.
- Acceptance: In cloud-hosted contexts, archive and list work.

## Implementation notes (practical)
- Reuse your existing same-workbook CF logic as-is for intra-workbook checks; cross-workbook uses the diff engine + batch formatting.
- Use document settings to remember: last overlay rect, applied addresses per sheet, last archive folder name, feature flags.
- Favor `RangeAreas` and single color sets per category to minimize traffic.
- Treat dates as numbers for equality (Excel serials); compare strings trimmed; compare formulas by normalized text.
- Skip hidden sheets in MVP; add a setting to include them later.

## Edge cases to track (later)
- Renamed/moved sheets (name mismatch but similar content) — out of scope for MVP; can add heuristics later.
- Tables/pivots/charts/shapes/VBA — ignored in MVP.
- Number formats vs values — future enhancement.
- Errors (#DIV/0! etc.) — compare as values; formula-change still orange.

## Success checklist
- Archive snapshot (IndexedDB) — available and durable across reloads.
- Upload baseline (.xlsx) — selectable and parsed correctly.
- Diff summary — accurate counts by sheet and total.
- Lazy cell coloring — applied only on activation; Clear works.
- Sheet tab colors — reflect severity and can be reset.
- Archive download — user gets a viable copy; fallback path documented.

## Commit message suggestions
- chore(ui): add archive/compare controls and messaging
- feat(core): extract WorkbookModel from current workbook
- feat(snapshot): IndexedDB save/list/delete
- feat(import): upload and parse .xlsx via xlsx
- feat(diff): cross-workbook diff engine (values/formulas)
- feat(format): batch apply/clear with RangeAreas
- feat(lazy): apply diffs on active sheet only; add toggle
- feat(tabs): color worksheet tabs by status
- feat(archive): user-downloadable workbook copy
- perf(format): chunking and sync throttling
- docs(test): expand docs and unit tests
- feat(cloud): OneDrive/SharePoint archive (optional)
