# PROJECT_SPECIFICATION

Updated: 2026-08-28

## Product

Excel Tab B24 is a static, serverless Bitrix24 Marketplace application embedded as a custom tab in a CRM deal card.

## Core Capabilities

- Shows a spreadsheet-like grid inside the deal card.
- Starts with 9 rows and 7 columns.
- The status under the title names the active sheet clearly: the deal sheet shows `Таблица сделки "..." из воронки ...`, and the funnel sheet shows `Общая таблица сделок из воронки ...`.
- Lets the user add rows with the plus button below the table.
- Lets the user add columns with the plus button to the right of the table. Adding a column preserves existing cell values, field bindings, and formatting, including rows that were previously shorter than the visible table width.
- Each cell can be edited manually.
- Cells can be selected individually, as ranges, by row, by column, or by the whole table.
- Cells can be selected by dragging across the grid while holding the left mouse button.
- Column widths can be changed manually by dragging column header borders.
- Row heights can be changed manually by dragging row header borders.
- When a cell is active, a field picker button is available.
- The hidden field picker control does not reserve permanent cell width; normal cell text uses the full column width and the picker appears as an overlay only when needed.
- The field picker lists CRM deal fields and inserts the selected field value into the current cell.
- The field picker can be closed explicitly with the close button, Escape, outside click, or repeated click on the same picker button.
- When cells are selected, the user can enable text wrapping for selected cells or auto-fit selected columns to their contents.
- When cells are selected, the user can clear them with `Удалить`. Clearing removes cell values, CRM field bindings, fill color, font weight, and wrapping. Selections larger than 18 cells require confirmation.
- The auto-fit action is always available. Without a selection it auto-fits all columns; with a selection it auto-fits only selected columns. It measures the actual grid value regardless of whether the value was typed manually or inserted/refreshed from a CRM deal field.
- Cells can contain formulas such as `=E4+B4`; formulas support cell references, numbers, parentheses, and `+`, `-`, `*`, `/`.
- After typing `=`, clicking another cell inserts that cell reference into the active formula. Repeated clicks append references with `+` by default, while clicks after a typed operator append the next reference after that operator.
- When a cell enters formula mode by starting with `=`, a suggestion popover opens under the cell with up to 5 recently used formulas. Selecting a suggestion writes it into the active cell.
- Formula cells store the original formula and display the calculated value. When the cell is focused, the editable formula is shown.
- Pressing Enter in a cell saves the current value or formula, exits edit mode, and shows the calculated formula result when applicable.
- When a cell is selected, the toolbar shows `Добавить формулу`. It opens a modal where the user can choose or delete a saved formula, save a new formula in the right-side formula entry panel, apply the selected formula to the active selected cell, or cancel. Formula entry blocks non-English characters, automatically uppercases English letters, and explains that formulas must use English letters, numbers, and formula symbols.
- The toolbar is visually compact, grouped with labeled sections, and uses a short `⇔ Подогнать` action for auto-fitting selected columns and rows.
- The toolbar has a separate `Бэкап` group with undo, redo, and delete controls. Undo and redo keep up to 15 in-memory table states per sheet/deal.
- When a formula cell is selected and the user Ctrl-selects additional cells, the formula is copied into those cells with relative row and column reference shifts, including row and column header Ctrl-selection.
- When cells are selected, the user can apply fill color, bold text, italic text, and font size presets: 11 pt, 13 pt, 15 pt, and 18 pt. The default table font size is 13 pt, and the 13 pt toolbar option explicitly applies to multi-cell selections.
- Deal reference fields are displayed as human-readable values where possible: users, contact, company, category, and stage.
- Cells filled from a CRM deal field keep a local field binding and refresh from the current deal each time the tab opens or fields are reloaded.
- The user can export the filled table area to an Excel-compatible `.xls` file trimmed to the last filled row and column, including exported column widths, 1.5 pt borders, fill color, font weight, italic style, font size, and alignment where applied. Formula cells are exported as Excel formulas, plain numeric cells are exported as numeric values instead of text, and cells without explicit fill color do not receive forced white fill.
- When the user clicks `Обновить поля`, refreshed field-bound values are loaded, the current table is compacted to the last filled row and column, and columns/rows are auto-fit while keeping at least the default 9 rows and 7 columns.
- The app iframe itself must not show or require any outer application scroll; scrolling is allowed only inside the table area when rows or columns exceed the visible grid.
- The initial visible table area fits the header row plus all 9 default rows above the horizontal scrollbar.
- Shift+Enter remains available for multiline wrapped text.

## Runtime Model

- Marketplace package is browser-only and serverless.
- Bitrix24 REST calls are made through `BX24.callMethod`.
- Initial REST methods used:
  - `placement.bind`
  - `placement.unbind`
  - `crm.deal.fields`
  - `crm.deal.get`
  - `crm.contact.get`
  - `crm.company.get`
  - `crm.dealcategory.list`
  - `crm.status.list`
  - `user.get`
- Current local persistence uses `localStorage`.
- Grid data is separated per deal when the deal ID is detected. The storage key format is `excel-tab-b24-grid-deal-v1-{dealId}`.
- Per-deal local state also stores wrapped cells and custom column widths.
- Per-deal local state also stores custom row heights.
- Per-deal local state also stores cell fill color, font weight, italic style, and font size formatting.
- Per-deal local state stores field bindings for cells filled from CRM deal fields.
- Saved reusable formulas are stored in browser `localStorage` under `excel-tab-b24-saved-formulas-v1`.
- Recently used formulas are stored in browser `localStorage` under `excel-tab-b24-recent-formulas-v1`, deduplicated with the newest formula first and limited to 5 items.
- If the app is opened outside a detected deal card, it falls back to the local development key `excel-tab-b24-grid-v1`.

## Required Bitrix24 Permissions

Before each Marketplace app upload, Codex must explicitly tell the user which Bitrix24 permissions are required for the current app version to work correctly. The permissions must be written as a list.

For the current version, required permissions are:

- CRM (CRM)
- Placement (Встраивание приложений)
- `user_basic`

Do not request full `user` or `user.userfield` for the current version.

## Packaging

- Use `npm run package` to create a new versioned zip in `dist app B24 zip/`.
- Archive names must use `Excel Tab B24 v.N.zip`.
- Old archives must not be overwritten.
- Marketplace zip includes only runtime files:
  - `install.html`
  - `install.js`
  - `install.css`
  - `index.html`
  - `app.js`
  - `style.css`

## Test Coverage

- Unit tests cover grid creation, row/column growth, column names, formula parsing/evaluation/reference shifting, saved and recent formula helpers, cell clearing helpers, deal ID extraction, per-deal storage keys, reference value formatting, selection helper behavior, arithmetic helpers, field-bound cell refresh, sheet state persistence, cell formatting persistence, and Excel export trimming/escaping/styling.

## Stage 2 Scope

- Add mouse drag selection for selecting multiple cells while holding the left mouse button.
- Add manual column width resizing by dragging column borders.
- Add manual row height resizing by dragging row borders.
- Persist custom row heights together with the current sheet state.
- Verify drag selection and resizing inside the Bitrix24 iframe, including interactions with cell editing, formula reference picking, and the CRM field picker.

## History

- 2026-08-18: Project initialized from an empty GitHub repository. Added first serverless Marketplace scaffold, documentation, lint/test scripts, and versioned zip builder.
- 2026-08-18: Improved Bitrix24 deal ID detection from placement context and clarified that `user.userfield` is not required for the current app version.
- 2026-08-19: Added per-deal grid storage keys, explicit selected-cell state, and guarded field insertion into the selected cell only.
- 2026-08-19: Added explicit field picker close controls, reference field display values, and idempotent placement rebinding during install.
- 2026-08-19: Added multi-cell selection, select-filled-cells action, wrapped text cells, and auto-fit column width action.
- 2026-08-19: Added automatic refresh for field-bound cells and Excel-compatible export for the filled table area.
- 2026-08-21: Locked the app shell to the Bitrix24 iframe height so only table rows/columns scroll, and made Enter commit the active cell value.
- 2026-08-20: Added always-available column auto-fit, selected-cell arithmetic, fill colors, font weight controls, formatting persistence, and styled Excel export.
- 2026-08-20: Changed the default grid to 9 rows by 7 columns, made the select action target the whole table, removed calculation status hints, and added Excel-like formulas with relative Ctrl-fill behavior.
- 2026-08-25: Removed the outer app scrollbar while preserving table scrollbars and made Enter save formulas/cell edits and leave edit mode.
- 2026-08-27: Added saved formulas modal and restored the default grid to 9 rows by 7 columns.
- 2026-08-27: Added saved formula deletion, English-only formula input guidance, formula list scrolling, and selected-cell clearing with confirmation for large selections.
- 2026-08-27: Clarified active sheet status text and renamed bottom sheet switch buttons.
- 2026-08-27: Added recent formula suggestions under the active cell when formula input starts with `=`.
- 2026-08-27: Prepared Marketplace v.17 with safer column growth, explicit 13 pt multi-cell font sizing, manual compacting on field reload, and formula-preserving `.xls` export.
- 2026-08-27: Prepared Marketplace v.18 with a taller default table viewport so the 9th default row is not hidden by the horizontal scrollbar.
- 2026-08-27: Prepared Marketplace v.19 so the hidden field picker no longer clips visible cell text after columns are added.
- 2026-08-28: Added `RELEASE_REPORT_v19.md` with management-facing v.19 summary and stage 2 implementation scope.
- 2026-08-28: Prepared Marketplace v.20 so exported `.xls` formulas can calculate against numeric cells instead of text-formatted numbers.
- 2026-08-28: Prepared Marketplace v.21 so Excel export uses 1.5 pt borders and does not force a white table area when cells have no fill color.
- 2026-08-28: Prepared Marketplace v.22 so Excel export preserves application column widths and stable row height instead of opening with squeezed columns.
- 2026-08-28: Prepared Marketplace v.23 with grouped alignment toolbar controls, short hover hints, multi-cell alignment formatting, and alignment export.
- 2026-08-28: Prepared Marketplace v.24 with visible vertical alignment in the app, alignment toolbar placement after text wrapping, and auto-fit width measurement by the longest line inside multiline cells.
- 2026-08-28: Prepared Marketplace v.25 with the version number visible in the app title and Bitrix24 iframe resizing so the app has no empty bottom area or outer scroll.
- 2026-08-28: Prepared Marketplace v.26 with mouse drag selection, manual column resizing, and persisted manual row heights.
- 2026-08-28: Prepared Marketplace v.27 with compact grouped toolbar controls and row-height auto-fit in the `⇔ Подогнать` and field refresh flows.
- 2026-08-28: Prepared Marketplace v.29 with a separate Backup toolbar group, 15-step per-sheet undo/redo, and matched export/reload button sizing.
