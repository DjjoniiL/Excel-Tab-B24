# NEXT_SESSION

Updated: 2026-08-27

## Current State

- Repository `DjjoniiL/Excel-Tab-B24` is connected as `origin`.
- Serverless Bitrix24 Marketplace app scaffold is present and actively developed.
- Runtime files are plain browser assets: `install.html`, `install.js`, `install.css`, `index.html`, `app.js`, `style.css`.
- Repository overview is documented in `README.md`.
- No backend, OAuth secrets, API keys, `.env`, documentation, tests, or `node_modules` are included in Marketplace zip output.
- Marketplace zip archives are created in `dist app B24 zip/`.
- Marketplace zip names use `Excel Tab B24 v.N.zip`, where `N` is the next available version number.
- Latest local Marketplace zip after this iteration: `dist app B24 zip/Excel Tab B24 v.15.zip`.
- A compact next-session prompt is stored in `PROMT_NEXT_Ses.txt`.
- The deal tab shows a 9x7 spreadsheet-like grid by default.
- The default visible table window fits the column header plus all 9 default rows before added rows require table scrolling.
- Users can add rows and columns while keeping scrolling inside the table area.
- The app iframe itself should not show a separate right-side page scrollbar.
- Users can select the whole table, individual cells, ranges, rows, and columns.
- Users can enable wrapping, auto-fit columns, export to Excel-compatible `.xls`, apply fill colors, and set font weight.
- Users can clear selected cells together with data, field bindings, fill color, font weight, and wrapping; selections larger than 18 cells ask for confirmation.
- Formula cells support references such as `=E4+B4`, clicked cell references after typing `=`, relative Ctrl-fill behavior, and Enter-to-save edit mode.
- Users can save reusable formulas in a modal, delete formulas from the saved list, and apply a selected saved formula to the active selected cell. Formula input in the modal and in cells removes non-English characters, uppercases English letters, and shows a user-facing explanation in the modal.

## Next Work

- Verify on the portal that the outer app scrollbar is gone while the table scrollbars remain for added rows and columns.
- Verify Enter saves a typed formula and exits formula edit mode.
- Verify formula reference picking after typing `=` works inside the Bitrix24 iframe.
- Verify relative Ctrl-fill for formulas across rows and columns.
- Verify saved formulas can be saved, selected, cancelled, and applied to the intended selected cell in the Bitrix24 iframe.
- Verify saved formula deletion, formula-list scrolling, and English-only formula input guidance in the Bitrix24 iframe.
- Verify `Удалить` clears selected cells and that deleting more than 18 selected cells requires confirmation.
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
- User basic / `user_basic` (Пользователи, базовый)

Do not request full `user` or `user.userfield` for the current version.
