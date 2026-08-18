# NEXT_SESSION

Updated: 2026-08-18

## Current State

- Repository `DjjoniiL/Excel-Tab-B24` is connected as `origin`.
- Initial serverless Bitrix24 Marketplace app scaffold is present.
- Runtime files are plain browser assets: `install.html`, `install.js`, `install.css`, `index.html`, `app.js`, `style.css`.
- No backend, OAuth secrets, API keys, `.env`, documentation, tests, or `node_modules` are included in Marketplace zip output.
- Marketplace zip archives are created in `dist app B24 zip/`.
- Marketplace zip names use `Excel Tab B24 v.N.zip`, where `N` is the next available version number.

## Next Work

- Verify the exact Bitrix24 placement code on a test portal. Current installer binds `CRM_DEAL_DETAIL_TAB`.
- Test inside a real deal card iframe to confirm `PLACEMENT_OPTIONS` shape and deal ID extraction.
- Decide whether table data should be per-deal, per-user, or shared through Bitrix24 app options. Current persistence is browser `localStorage`.
- Add UX polish after first portal test: keyboard navigation, row/column resizing, and clearer selected-cell state.

## Required Cycle

1. Run `git status --short`.
2. Make scoped changes.
3. Run `npm run lint`.
4. Run `npm test`.
5. Commit changes.
6. Ask for permission before `git push`.

## Marketplace Upload Permissions

Before uploading each app version, announce and verify:

- CRM (CRM)
- User (Пользователи)
- Placement (Встраивание приложений)
