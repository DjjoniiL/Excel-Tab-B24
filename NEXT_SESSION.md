# NEXT_SESSION

Updated: 2026-08-25

## Current State

- Repository `DjjoniiL/Excel-Tab-B24` is connected as `origin`.
- Serverless Bitrix24 Marketplace app scaffold is present and actively developed.
- Runtime files are plain browser assets: `install.html`, `install.js`, `install.css`, `index.html`, `app.js`, `style.css`.
- Repository overview is documented in `README.md`.
- No backend, OAuth secrets, API keys, `.env`, documentation, tests, or `node_modules` are included in Marketplace zip output.
- Marketplace zip archives are created in `dist app B24 zip/`.
- Marketplace zip names use `Excel Tab B24 v.N.zip`, where `N` is the next available version number.
- Latest local Marketplace zip after this iteration: `dist app B24 zip/Excel Tab B24 v.12.zip`.
- A compact next-session prompt is stored in `PROMT_NEXT_Ses.txt`.
- The deal tab shows a 9x7 spreadsheet-like grid by default.
- Users can add rows and columns while keeping scrolling inside the table area.
- The app iframe itself should not show a separate right-side page scrollbar.
- Users can select the whole table, individual cells, ranges, rows, and columns.
- Users can enable wrapping, auto-fit columns, export to Excel-compatible `.xls`, apply fill colors, and set font weight.
- Formula cells support references such as `=E4+B4`, clicked cell references after typing `=`, relative Ctrl-fill behavior, and Enter-to-save edit mode.

## Next Work

- Verify on the portal that the outer app scrollbar is gone while the table scrollbars remain for added rows and columns.
- Verify Enter saves a typed formula and exits formula edit mode.
- Verify formula reference picking after typing `=` works inside the Bitrix24 iframe.
- Verify relative Ctrl-fill for formulas across rows and columns.
- Verify field-bound cells refresh automatically when a deal field changes.
- Verify Excel export opens correctly in desktop Excel and includes calculated formula values.
- Add further UX polish after portal test: keyboard navigation and manual row/column resizing.

## Required Cycle

1. Run `git status --short`.
2. Make scoped changes.
3. Run `npm run lint`.
4. Run `npm test`.
5. Commit changes.
6. Ask for permission before `git push`, unless the user explicitly requested push in the current iteration.

## Marketplace Upload Permissions

Before each Marketplace app upload, tell the user which Bitrix24 permissions are required for the current app version to work correctly. Always write the permissions as a list.

For the current version, required permissions are:

- CRM (CRM)
- Placement (Встраивание приложений)
- User (Пользователи)

Do not request `user.userfield` for the current version.
