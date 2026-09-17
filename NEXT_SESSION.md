# NEXT_SESSION

Обновлено: 2026-09-16

## Финальное состояние

Проект `DjjoniiL/Excel-Tab-B24` доведен до релизной версии:

```text
Excel таблицы в сделке и экспорт данных из CRM v.47
```

v47 зафиксирована как релизная версия с добавленным встроенным окном помощи, кнопкой `Помощь` и виджетом открытой линии.

## Локальный путь

```text
G:\AI Project B24\Excel Tab B24
```

## GitHub

```text
https://github.com/DjjoniiL/Excel-Tab-B24
```

## Финальный пакет

```text
dist app B24 zip/excel-tab-b24-v47.zip
```

Скрипт `npm run package` не перезаписывает существующий zip. Для повторной сборки v47 нужно вручную удалить старый архив или поднять номер версии в `tools/build-marketplace-zip.ps1`.

## Актуальные файлы

- `README.md` - публичное описание релиза и инструкция.
- `PROJECT_SPECIFICATION.md` - техническая спецификация релиза.
- `DESIGN_GUIDE.md` - правила интерфейса.
- `RELEASE_REPORT_v47.md` - итоговый отчет по релизу.
- `tests/app.test.js` - unit-тесты ключевой логики.

## Runtime-файлы

- `install.html`
- `install.js`
- `install.css`
- `index.html`
- `app.js`
- `style.css`

В Marketplace zip должны попадать только эти runtime-файлы.

## Ключевые правила релиза

- Приложение остается serverless и browser-only.
- Данные хранятся в Bitrix24 entity storage `exctabb24`; `localStorage` используется как кеш/fallback.
- Индикатор количества полей сделки удален из UI полностью.
- Новая сделка должна открывать пустой лист в группе `Листы сделки`.
- До определения ID сделки используется ключ `excel-tab-b24-grid-pending-deal-v1`.
- ID из текущего URL/referrer имеет приоритет над устаревшими placement-параметрами.
- Общие листы остаются общими для воронки и хранятся по `excel-tab-b24-grid-funnel-v1-{categoryId}`.
- Листы сделки хранятся по `excel-tab-b24-grid-deal-v1-{dealId}`.
- Максимальный размер листа: 3000x3000.
- Максимум листов в каждой группе: 7.

## Проверки перед пушем

```bash
npm run lint
npm test
```

Перед коммитом проверить `git status --short` и не добавлять локальные/служебные untracked-файлы без отдельного запроса пользователя.

Текущие untracked-файлы, которые не нужно добавлять автоматически:

- `icon-preview.html`
- `marketplace-assets/`
- `Скриншоты/`

## Права Bitrix24

- CRM
- Placement / Встраивание приложений
- Entity / Хранилище данных приложения
- `user_brief`

Полные `user`, `user_basic` и `user.userfield` не нужны.

## REST-методы

- `placement.unbind`
- `placement.bind`
- `entity.add`
- `entity.update`
- `entity.item.property.add`
- `entity.item.get`
- `entity.item.add`
- `entity.item.update`
- `crm.deal.fields`
- `crm.deal.get`
- `crm.contact.get`
- `crm.company.get`
- `crm.dealcategory.list`
- `crm.status.list`
- `user.get`
