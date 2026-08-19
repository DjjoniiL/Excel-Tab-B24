# NEXT_SESSION

Updated: 2026-08-18

## Current State

- Repository `DjjoniiL/Excel-Tab-B24` is connected as `origin`.
- Initial serverless Bitrix24 Marketplace app scaffold is present.
- Runtime files are plain browser assets: `install.html`, `install.js`, `install.css`, `index.html`, `app.js`, `style.css`.
- No backend, OAuth secrets, API keys, `.env`, documentation, tests, or `node_modules` are included in Marketplace zip output.
- Marketplace zip archives are created in `dist app B24 zip/`.
- Marketplace zip names use `Excel Tab B24 v.N.zip`, where `N` is the next available version number.
- A compact next-session prompt is stored in `PROMT_NEXT_Ses.txt`.

## Next Work

- Verify the exact Bitrix24 placement code on a test portal. Current installer binds `CRM_DEAL_DETAIL_TAB`.
- Test inside a real deal card iframe to confirm `PLACEMENT_OPTIONS` shape and deal ID extraction.
- Verify field reference display values on a real portal: users, contact, company, category, and stage.
- Verify that field-bound cells refresh automatically when a deal field changes.
- Verify Excel export opens correctly in desktop Excel.
- Verify that repeated app installation leaves one deal tab after `placement.unbind` + `placement.bind`.
- Add further UX polish after portal test: keyboard navigation and manual row/column resizing.

## Required Cycle

1. Run `git status --short`.
2. Make scoped changes.
3. Run `npm run lint`.
4. Run `npm test`.
5. Commit changes.
6. Ask for permission before `git push`.

## Marketplace Upload Permissions

Before each Marketplace app upload, tell the user which Bitrix24 permissions are required for the current app version to work correctly. Always write the permissions as a list.

For the current version, required permissions are:

- CRM (CRM)
- Placement (Встраивание приложений)
- User (Пользователи)

Do not request `user.userfield` for the current version.
