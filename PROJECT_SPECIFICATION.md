# PROJECT_SPECIFICATION

Updated: 2026-08-18

## Product

Excel Tab B24 is a static, serverless Bitrix24 Marketplace application embedded as a custom tab in a CRM deal card.

## Core Capabilities

- Shows a spreadsheet-like grid inside the deal card.
- Starts with 12 rows and 8 columns.
- Lets the user add rows with the plus button below the table.
- Lets the user add columns with the plus button to the right of the table.
- Each cell can be edited manually.
- When a cell is active, a field picker button is available.
- The field picker lists CRM deal fields and inserts the selected field value into the current cell.

## Runtime Model

- Marketplace package is browser-only and serverless.
- Bitrix24 REST calls are made through `BX24.callMethod`.
- Initial REST methods used:
  - `placement.bind`
  - `crm.deal.fields`
  - `crm.deal.get`
- Current local persistence uses `localStorage`.
- Required Bitrix24 permissions before Marketplace upload: CRM and placement/application embedding permissions.

## Packaging

- Use `npm run package` to create a new versioned zip in `dist/`.
- Old archives must not be overwritten.
- Marketplace zip includes only runtime files:
  - `install.html`
  - `install.js`
  - `install.css`
  - `index.html`
  - `app.js`
  - `style.css`

## History

- 2026-08-18: Project initialized from an empty GitHub repository. Added first serverless Marketplace scaffold, documentation, lint/test scripts, and versioned zip builder.
