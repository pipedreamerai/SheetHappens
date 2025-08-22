## SheetHappens - Excel Workbook Versioning

Compare your current Excel workbook against a baseline (another .xlsx or a saved snapshot) and visualize differences directly in the workbook. All data stays local; no cloud integration and no file export.

### Features

- Cross-workbook compare: current workbook vs uploaded file or local snapshot
- Snapshot archive: save/load/delete baselines in IndexedDB
- In-sheet highlights (conditional formats):
  - Green: added
  - Red: removed
  - Yellow: value changed
  - Orange: formula changed
- Sheet tab colors reflect severity per sheet (red > orange > yellow > green)
- Auto apply-per-sheet: formatting is applied when you activate a sheet
- Selection callout: select a changed cell to see “New / Old” values
- One-click cleanup: Stop Diff removes all highlights and resets tab colors
- Revert cells with a one-click "Revert" button
- Operates only on cell values and NOT on formatting

## Quickstart

### Requirements

- Excel Desktop (macOS or Windows)
- Excel 365
- Internet access to load the add-in assets

### Install (no dev required)

1. Download the production manifest: [Download `manifest.xml`](https://raw.githubusercontent.com/pipedreamerai/SheetHappens/main/manifest.xml)
2. macOS (Excel Desktop)
   - Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create the `wef` folder if it doesn’t exist), then restart Excel.
3. Windows or Excel 365
   - Sideload the add-in (Worked for me on Excel 365 by going to Add-ins -> More Add-ins -> My Add-ins -> Upload My Add-in and then selecting the manifest file)
4. Open any workbook. In the Home tab, you should see the command group “SheetHappens”. Click “SheetHappens” to open the task pane.

Notes

- The manifest points to a hosted, secure URL (`https://pipedreamerai.github.io/SheetHappens/addin/taskpane.html`). You do not need to run a server.
- If you previously had Excel open while copying the manifest, fully quit and relaunch Excel.

## Using the add-in

1. Take a snapshot (optional)
   - Click “Take Snapshot” to store the current workbook as a local baseline.
2. Choose a baseline
   - Upload: Click “Choose File”, pick a .xlsx, then select it under “Baseline (uploads)”.
   - Snapshot: Pick one under “Baseline (snapshots)”.
3. Start the comparison
   - Click “Start Diff”. Sheet tabs are colored by severity; highlights appear as you activate sheets.
4. Review changes
   - Switch between sheets. Highlights are applied automatically when a sheet becomes active.
   - Select a highlighted cell to see a “New / Old” callout.
5. Stop the comparison
   - Click “Stop Diff” to remove all highlights and reset tab colors.
6. Optional: Revert a single cell
   - Select a changed cell and click “Revert Selection” to restore it to the baseline (added/removed/formula changes).

### Color semantics

- Green: present now, blank in baseline (added)
- Red: blank now, present in baseline (removed)
- Orange: formula changed (FORMULATEXT differs)
- Yellow: value changed (same formula text)

## Notes and limitations

- Local-only: snapshots are stored in your browser’s IndexedDB; no OneDrive/SharePoint
- Visible sheets only (MVP)
- Tables/pivots/charts/shapes/VBA are ignored
- Dates are compared by numeric value (Excel serials)
- Strings are compared trimmed; formulas compared by normalized text

## Troubleshooting

- “Select a baseline first”: upload a file or pick a snapshot, then click Start Diff
- “Failed to parse upload”: ensure a valid .xlsx
- If Excel was already running, close and try `npm run dev` again

## Development

- Build (production): `npm run build`
- Key files
  - `src/core/model.js`: build WorkbookModel from the active workbook
  - `src/core/import-xlsx.js`: parse uploaded .xlsx into a model
  - `src/core/diff.js`: pure diff engine
  - `src/core/snapshot.js`: IndexedDB save/load/delete
  - `src/taskpane/taskpane.js`: UI wiring and formatting

## FAQ

- Be the first to ask questions!
