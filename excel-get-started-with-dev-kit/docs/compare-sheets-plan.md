## Excel Add-in: Sheet-to-Sheet Compare — Implementation Plan

This document outlines how to implement an Excel task pane add-in that compares two worksheets cell-by-cell and highlights differences according to the rules below. This is a plan only; no implementation is included.

### Requirements (from user)
- Compare two sheets at the same row/column coordinates.
- If two cells are exactly the same (including both value type and content), do nothing.
- If two cells have different hardcoded values (neither is a formula), highlight orange.
- If two cells have the same formula text but different resulting values, highlight yellow.
- If the source sheet cell has content and the second sheet cell is blank, highlight green.
- If the source sheet cell is blank and the second sheet cell has content, highlight red.
- If two cells have different number formats that display the same value, highlight light blue.

Notes and clarifications to adopt for the first version:
- “Content” means a formula or a non-empty value. Blank means no formula and value is empty or null.
- “Exactly the same” means:
  - If both are formulas: formula strings are equal (case-insensitive, trimmed) AND their calculated values are equal (strict value equality with type-aware comparison; see Equality below).
  - If both are non-formulas: values equal (type-aware).
- Format-only differences (e.g., same value displayed differently due to `numberFormat`) are light blue.

### Architecture for maintainability
- Core comparator module (pure, framework-agnostic):
  - Input: two matrices of CellSnapshot plus options.
  - Output: a matrix of CompareOutcome enums and color buckets (addresses per color).
  - Unit-testable without Office.js by using plain JS arrays.
- Office.js adapter:
  - Reads ranges, maps to CellSnapshot, writes highlights using Range/RangeAreas.
- Task pane (UI):
  - Orchestrates selections, options, progress, and cancel.

### User Experience
- Entry points
  - Ribbon: “Compare Sheets” button under a custom group (e.g., “Versioning Tools”).
  - Task pane: Hosts the compare UI and progress.

- Task pane controls
  - Source sheet dropdown (required).
  - Second sheet dropdown (required).
  - Range mode: “Used range of both sheets” (default) or “Selected range”.
    - If Selected range: Use the active selection on the source sheet and project bounds to the same address on the second sheet. If the second sheet doesn’t have that range, still compare using the implied bounds, treating missing cells as blank.
  - Options (checkboxes):
    - Ignore case for text comparison (default: off).
    - Trim whitespace before compare (default: on).
    - Numeric tolerance (optional, default off): equality within epsilon (e.g., 1e-10) to avoid floating artifacts.
  - Compare by displayed text instead of raw values (default: off). When on, uses Range.text for equality.
  - Coerce numeric strings ("1" vs 1) before numeric compare (default: off).
  - Actions:
    - “Run Compare” (primary).
    - “Clear Highlights”.
  - Status/progress area: rows processed, estimated time, cancel button.

### Highlight colors
- Orange (different hardcoded values): `#FFA500` (or Excel theme orange).
- Yellow (same formula, different results): `#FFF2CC` (light yellow is easier on the eyes than pure `#FFFF00`).
- Green (source has content, second is blank): `#C6EFCE` (Excel’s default light green).
- Red (source blank, second has content): `#FFC7CE` (Excel’s default light red).
 - Light blue (format-only differences): `#CFE2F3` (Excel’s default light blue).

These are fill colors applied to the compared cells on both sheets.

### Data/Detection model (Office.js)
- Read in bulk to minimize round-trips:
  - For the comparison range on both sheets, load: `values`, `formulas`, `numberFormat`, `text`, and `valueTypes`.
  - Calculate max bounds to compare: use the union of used ranges (or the selected range option).
- Formula detection:
  - A cell is “formula” if its `formulas` entry is a non-empty string starting with `=`.
  - “Same formula” means the exact formula text matches after: trimming, normalizing case, and optionally normalizing localized vs invariant forms if we use `formulaLocal`. Prefer invariant `formulas`.
- Blank detection:
  - formula: absent.
  - value: `null` or `""` (after trim if text) is blank.
 - Number format difference detection:
   - If values are equal (per Equality rules or displayed-text mode) but `numberFormat` differs, treat as light blue.

### Equality rules
- Text: equal if strings match; optionally apply case-insensitive and trim if options selected.
- Numbers: equal if exactly equal, or within epsilon if tolerance is enabled.
- Booleans: strict equality.
- Errors (e.g., `#N/A`): treat as string-labeled values; equal if the same error code.
- Dates: Office.js returns dates as numbers with date formats; treat them as numbers unless we detect date category and compare numeric serials.
 - Displayed text mode: if enabled, compare `text` values (post-format) rather than raw `values`.
 - Numeric string coercion: if enabled, coerce parseable strings to numbers before numeric comparison.

### Rule precedence and conflict resolution
Apply rules in this order per cell pair:
1) Source has content, second is blank => Green.
2) Source is blank, second has content => Red.
3) Both formulas and normalized formulas equal AND resulting values differ => Yellow.
4) Values equal but numberFormat differs => Light blue (format-only).
5) Neither is a formula AND values differ => Orange.
6) Else => No highlight.

### Algorithm (high level)
1. Input validation: ensure two distinct sheets selected; compute comparison range.
2. Read both ranges in one batch: formulas and values arrays for source and second.
3. Iterate row-major once, evaluating each cell pair:
   - Compute `srcIsBlank`, `dstIsBlank`, `srcIsFormula`, `dstIsFormula`.
  - Apply rules in the precedence order above, including the format-only difference rule.
4. Collect address buckets per color (to reduce formatting calls): build sparse areas by rows/columns or contiguous ranges where possible.
5. Apply fills in batches per color on both sheets.
6. Report summary counts; keep a “Clear Highlights” feature that resets fills for the compared range(s).

### Performance considerations
- Always read and write in bulk.
- For very large ranges (>200k cells), process by chunks (e.g., by 1k–5k rows at a time) to avoid payload overflows and keep UI responsive.
- Use a cancel token flag in the task pane (cooperative cancellation between chunks).
- Avoid setting fill per-cell; group addresses into multi-range addresses or use `getRanges()` with a list of address strings per color.
 - Use RangeAreas when available (ExcelApi >= 1.9); otherwise coalesce contiguous ranges by rows/cols.

### Edge cases and behaviors
- Merged cells: treat based on top-left cell; warn users that results on merged areas may vary.
- Spilled arrays: Office.js may mark only the origin as formula; consider the displayed values in spill area as non-formulas — this means yellow may not appear in spillage unless we do additional checks. First version: no special handling.
- Different formulas with same result: no highlight (by spec omission).
- Dynamic arrays and volatile functions: values may change between reads; comparison runs on a single snapshot read.
- Hidden rows/columns: still compared; highlighting applies even if hidden.
- Protection: if sheets are protected and formatting is blocked, show an actionable error.
- External links and errors: treat as values; equal if both error codes equal.
 - Text vs number that display the same (e.g., "1" vs 1): equal only when displayed-text mode or numeric coercion is enabled; otherwise considered different.

### Manifest and UI wiring (planned changes)
- Add a ribbon button "Compare Sheets" mapping to a task pane command.
- Task pane HTML adds controls listed above.
- Permissions: require `ReadWriteDocument` (formatting writes) in the manifest.
- Requirement sets: target ExcelApi 1.9+ (RangeAreas), with graceful fallback.

### Implementation steps (tracked tasks)
1) Task pane UI
   - Add dropdowns for Source and Second sheets, range mode controls, options, actions, and status area.
2) Sheet discovery and selection
   - Enumerate worksheets for dropdowns; handle active sheet preselection.
3) Range determination
   - Implement Used-range union and Selected-range path.
4) Bulk read
  - Load values, formulas, numberFormat, text, and valueTypes for both ranges.
5) Comparison engine
  - Implement type-aware equality, blank detection, format-difference detection, and rule ordering.
  - Build comparator as a pure module with a clear input/output contract and unit tests.
6) Batching and formatting
   - Group addresses per color and apply fills in bulk on both sheets.
7) Clear highlights
   - Provide a reset action to clear fills in the last compared ranges.
8) Progress, cancellation, and error handling
   - Update UI during chunked processing; support cancel.
9) Testing and validation
   - Unit tests for comparator (happy path + edge cases).
   - Manual scenarios and large-range stress checks.

### Commit-by-commit plan (each step is user-testable)
1) Wire up "Compare Sheets" ribbon button and empty task pane (done)
  - Visible change: New ribbon button opens a task pane with a placeholder title.
  - Test: Click button, pane opens without errors.
2) Basic task pane UI shell (done)
  - Add dropdowns for Source/Second sheets, range mode toggles, and disabled "Run Compare" button.
  - Test: UI renders; Run button is disabled until valid selections.
3) Populate worksheet dropdowns (done)
  - Enumerate workbook sheets and populate both dropdowns; prevent picking the same sheet twice.
  - Test: Sheets appear; selecting the same sheet shows a validation message.
5) Clear Highlights (in selection)
  - Implement a working "Clear Highlights" that clears fills in the computed range(s).
  - Test: Manually fill some cells, click Clear, fills are removed.
6) Bulk read plumbing (dry run)
  - Read values, formulas, numberFormat, text, valueTypes for the computed range; do not format yet.
  - Show counts in pane (cells, formulas, blanks) and enable a "Dry run" toggle.
  - Test: Run shows stats; Dry run produces no formatting.
7) Presence/absence rules (Green/Red)
  - Implement green/red for src-only and dst-only content; respect Selected/Used range and chunking.
  - Test: Differences highlight green/red; progress bar advances; Cancel stops cleanly.
8) Hardcoded value differences (Orange)
  - Implement non-formula value comparison with options (ignore case, trim, numeric tolerance).
  - Test: Create mismatched values; toggling options changes results as expected.
9) Same formula, different result (Yellow)
  - Implement formula normalization and result comparison; highlight yellow when formulas match but values differ.
  - Test: Same formula text producing different values highlights yellow.
10) Format-only differences (Light blue)
  - Implement detection where values equal but numberFormat differs; add displayed-text comparison option.
  - Test: 0.5 vs 50% lights blue; switching to displayed-text mode removes the difference if texts match.
11) RangeAreas batching + fallback
  - Apply highlights using RangeAreas when supported; fallback to coalesced contiguous ranges.
  - Test: Large selections apply quickly; verify behavior on platforms lacking RangeAreas.
12) Summary and navigation
  - Show per-color counts and allow clicking to jump to first difference on each sheet.
  - Test: Counts match visual highlights; navigation selects the expected cell.
13) Remember last compared range + Clear uses it
  - Cache last compared ranges; Clear Highlights clears those even across pane reload.
  - Test: Compare, reload pane, clear; fills are removed.
14) Options persistence and defaults
  - Persist user options in local storage; default source to active sheet and second to next sheet.
  - Test: Reload preserves choices; defaults feel sensible.
15) Robust errors and protection handling
  - Detect sheet protection/blocked formatting; show actionable error and skip formatting gracefully.
  - Test: Protect a sheet; run compare; UI shows message without crashing.
16) Verbose logs toggle
  - Add a dev-only toggle for logging timings and chunk metrics.
  - Test: Toggle prints logs in console without impacting user output.

### Minimal data model (for the core comparator)
- CellSnapshot
  - value: any | null
  - formula: string | null
  - numberFormat: string | null
  - text: string | null
  - valueType: string | null (e.g., "Double", "String", "Boolean", "Error")
- CompareOutcome: one of { none, orange_value_diff, blue_format_diff, yellow_formula_same_value_diff, green_src_only, red_dst_only }

### Testing plan (manual scenarios)
- Small 5x5 grid with mixtures of:
  - Identical values and identical formulas+values: no highlight.
  - Different hardcoded values: orange.
  - Same formulas but different current results (e.g., `=NOW()` vs delayed value, or same references but differing inputs): yellow.
  - Source content / Second blank: green.
  - Source blank / Second content: red.
  - Same resulting values but different number formats (e.g., 0.5 vs 50%): light blue.
- Date, boolean, number with rounding, and error values (#N/A) comparisons.
- Large used ranges (e.g., 5k x 50): performance sanity.
 - Text vs number comparisons under different options (displayed-text mode, numeric coercion).

### Telemetry / logging
- Optional console logging gated by a “Verbose logs” toggle in dev builds.

### Future enhancements (not in v1)
- Support a “Differences report” sheet with counts and addresses.
- Export result ranges to a new workbook.
- Option to highlight only on one of the two sheets.
- Option to treat different formulas with same result as an info highlight.
 - Deep format comparison (font, fill, borders) as additional categories.

### Acceptance criteria
- Correct application of the rules above (including format-only differences) across the entire chosen range.
- Reasonable performance for up to ~250k cells compared in under a minute on typical hardware.
- Non-destructive: only cell fill color is changed; Clear restores fills to no-fill for compared ranges.
- Clear user flow with basic validation and error messages.
 - Comparator covered by unit tests (blank/content, formula-same/result-diff, numeric tolerance, format-only diff, displayed-text mode).

---

Owner: TBD  •  Version: Draft 2  •  Date: 2025-08-15
