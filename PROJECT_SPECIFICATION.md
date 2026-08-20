# PROJECT_SPECIFICATION

Updated: 2026-08-21

## Product

Excel Tab B24 is a static, serverless Bitrix24 Marketplace application embedded as a custom tab in a CRM deal card.

## Core Capabilities

- Shows a spreadsheet-like grid inside the deal card.
- Starts with 12 rows and 8 columns.
- Lets the user add rows with the plus button below the table.
- Lets the user add columns with the plus button to the right of the table.
- Each cell can be edited manually.
- Cells can be selected individually, as ranges, by row, by column, or by all filled cells.
- When a cell is active, a field picker button is available.
- The field picker lists CRM deal fields and inserts the selected field value into the current cell.
- The field picker can be closed explicitly with the close button, Escape, outside click, or repeated click on the same picker button.
- When cells are selected, the user can enable text wrapping for selected cells or auto-fit selected columns to their contents.
- The auto-fit action is always available. Without a selection it auto-fits all columns; with a selection it auto-fits only selected columns. It measures the actual grid value regardless of whether the value was typed manually or inserted/refreshed from a CRM deal field.
- When cells are selected, the user can calculate addition, subtraction, multiplication, or division for numeric selected cell values and write the result into the active cell.
- When cells are selected, the user can apply a fill color and set font weight, including bold text.
- Deal reference fields are displayed as human-readable values where possible: users, contact, company, category, and stage.
- Cells filled from a CRM deal field keep a local field binding and refresh from the current deal each time the tab opens or fields are reloaded.
- The user can export the filled table area to an Excel-compatible `.xls` file trimmed to the last filled row and column, including fill color and font weight where applied.
- The deal tab iframe keeps scrolling inside the spreadsheet area only; the outer app shell does not scroll.
- Pressing Enter in the active cell saves the current value and closes the editor focus. Shift+Enter remains available for multiline wrapped text.

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
- Per-deal local state also stores cell fill color and font weight formatting.
- Per-deal local state stores field bindings for cells filled from CRM deal fields.
- If the app is opened outside a detected deal card, it falls back to the local development key `excel-tab-b24-grid-v1`.

## Required Bitrix24 Permissions

Before each Marketplace app upload, Codex must explicitly tell the user which Bitrix24 permissions are required for the current app version to work correctly. The permissions must be written as a list.

For the current version, required permissions are:

- CRM (CRM)
- Placement (Встраивание приложений)
- User (Пользователи)

Do not request `user.userfield` for the current version.

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

- Unit tests cover grid creation, row/column growth, column names, deal ID extraction, per-deal storage keys, reference value formatting, selection helper behavior, arithmetic helpers, field-bound cell refresh, sheet state persistence, cell formatting persistence, and Excel export trimming/escaping/styling.

## History

- 2026-08-18: Project initialized from an empty GitHub repository. Added first serverless Marketplace scaffold, documentation, lint/test scripts, and versioned zip builder.
- 2026-08-18: Improved Bitrix24 deal ID detection from placement context and clarified that `user.userfield` is not required for the current app version.
- 2026-08-19: Added per-deal grid storage keys, explicit selected-cell state, and guarded field insertion into the selected cell only.
- 2026-08-19: Added explicit field picker close controls, reference field display values, and idempotent placement rebinding during install.
- 2026-08-19: Added multi-cell selection, select-filled-cells action, wrapped text cells, and auto-fit column width action.
- 2026-08-19: Added automatic refresh for field-bound cells and Excel-compatible export for the filled table area.
- 2026-08-21: Locked the app shell to the Bitrix24 iframe height so only table rows/columns scroll, and made Enter commit the active cell value.
- 2026-08-20: Added always-available column auto-fit, selected-cell arithmetic, fill colors, font weight controls, formatting persistence, and styled Excel export.
