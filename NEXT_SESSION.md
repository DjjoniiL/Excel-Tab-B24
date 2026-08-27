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
- Latest local Marketplace zip after this iteration: `dist app B24 zip/Excel Tab B24 v.19.zip`.
- A compact next-session prompt is stored in `PROMT_NEXT_Ses.txt`.
- The deal tab shows a 9x7 spreadsheet-like grid by default.
- The default visible table window fits the column header plus all 9 default rows above the horizontal scrollbar before added rows require table scrolling.
- Active sheet status text is user-facing: deal mode says `Таблица сделки "..." из воронки ...`; funnel mode says `Общая таблица сделок из воронки ...`.
- Users can add rows and columns while keeping scrolling inside the table area.
- Adding columns preserves existing CRM field-bound cells, values, and formatting even if a saved table has uneven row widths.
- The app iframe itself should not show a separate right-side page scrollbar.
- Users can select the whole table, individual cells, ranges, rows, and columns.
- Hidden field picker buttons overlay the right edge of active cells instead of permanently taking width from every cell.
- Users can enable wrapping, auto-fit columns, export to Excel-compatible `.xls`, apply fill colors, bold text, italic text, and font size presets. Default table font size is 13 pt, and the 13 pt selector value applies explicitly to multi-cell selections.
- Users can clear selected cells together with data, field bindings, fill color, bold/italic/font-size formatting, and wrapping; selections larger than 18 cells ask for confirmation.
- Formula cells support references such as `=E4+B4`, clicked cell references after typing `=`, relative Ctrl-fill behavior, and Enter-to-save edit mode.
- Users can save reusable formulas in a modal, delete formulas from the saved list, and apply a selected saved formula to the active selected cell. Formula input in the modal and in cells removes non-English characters, uppercases English letters, and shows a user-facing explanation in the modal.
- Entering `=` in a cell shows up to 5 recently used formulas under that cell; selecting a suggestion writes it into the active cell.
- Formula cells export to `.xls` as Excel formulas, not only as calculated text.
- Clicking `Обновить поля` refreshes field-bound cells and compacts the current sheet to the last filled row and column, keeping at least the 9x7 default.

## Next Work

- Stage 1 portal check: verify adding columns no longer clears or shifts CRM field-bound cells.
- Stage 1 portal check: verify exported `.xls` opens in Excel with visible rows and formula cells remain formulas.
- Stage 1 portal check: verify `Обновить поля` refreshes CRM-bound values and compacts the table to the filled area, with a minimum of 9 rows and 7 columns.
- Stage 1 portal check: verify font size 11/13/15/18 pt applies to several selected cells.
- Stage 1 portal check: verify the 9th row is visible above the horizontal scrollbar after loading v.18.
- Stage 1 portal check: verify regular cell text is no longer clipped by the hidden field picker after adding columns in v.19.
- Stage 2 next work: add mouse drag selection with held left button.
- Stage 2 next work: add manual column width and row height resizing by dragging borders.
- Verify on the portal that the outer app scrollbar is gone while the table scrollbars remain for added rows and columns.
- Verify Enter saves a typed formula and exits formula edit mode.
- Verify formula reference picking after typing `=` works inside the Bitrix24 iframe.
- Verify relative Ctrl-fill for formulas across rows and columns.
- Verify saved formulas can be saved, selected, cancelled, and applied to the intended selected cell in the Bitrix24 iframe.
- Verify saved formula deletion, formula-list scrolling, and English-only formula input guidance in the Bitrix24 iframe.
- Verify `Удалить` clears selected cells and that deleting more than 18 selected cells requires confirmation.
- Verify active sheet status text and bottom switcher labels in both deal and funnel modes.
- Verify recent formula suggestions appear under the active cell after typing `=`, keep newest formulas first, and show no more than 5 items.
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
- `user_basic`

Do not request full `user` or `user.userfield` for the current version.
