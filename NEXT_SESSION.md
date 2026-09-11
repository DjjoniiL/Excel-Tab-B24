# NEXT_SESSION

Обновлено: 2026-09-10

## Финальное состояние

Проект `DjjoniiL/Excel-Tab-B24` доведен до `Excel таблица в сделке и экспорт MVP Final v.33`.

Текущий локальный путь:

```text
G:\AI Project B24\Excel Tab B24
```

GitHub:

```text
https://github.com/DjjoniiL/Excel-Tab-B24
```

Финальный Marketplace zip:

```text
dist app B24 zip/Excel таблица в сделке и экспорт MVP Final v.33.zip
```

## Что считать актуальным

- Приложение остается serverless и browser-only.
- Runtime-файлы: `install.html`, `install.js`, `install.css`, `index.html`, `app.js`, `style.css`.
- Видимое название: `Excel таблица в сделке и экспорт`.
- Номер версии: `v.33`, отображается мелким шрифтом в правом нижнем углу приложения.
- Под заголовком приложения больше нет описания сделки, общей таблицы или воронки.
- Нижний переключатель общей таблицы показывает `Таблица всех сделок` без названия воронки.
- Поля типа `date` выводятся как `16.09.2026`.
- Поля типа `datetime` выводятся как `03:00 (+3ч) 16.09.2026г.`.
- В Marketplace zip не входят документация, тесты, `.git`, `.env`, `node_modules`, локальные данные и служебные файлы.

## Документация

Перед будущими изменениями читать:

- `README.md`
- `PROJECT_SPECIFICATION.md`
- `DESIGN_GUIDE.md`
- верхнеуровневый `Bitrix24 CRM Marketplace App Logic.md`

## Права приложения

- CRM (CRM)
- Placement / Встраивание приложений
- `user_brief`

Полный `user`, `user_basic` и `user.userfield` не нужны.

## REST методы

- `placement.unbind`
- `placement.bind`
- `crm.deal.fields`
- `crm.deal.get`
- `crm.contact.get`
- `crm.company.get`
- `crm.dealcategory.list`
- `crm.status.list`
- `user.get`

## Контроль перед финалом

- `npm run lint`
- `npm test`
- `npm run package`
- Проверить, что zip создан с номером v.33.
- Проверить `git status --short`, чтобы случайные локальные файлы не попали в коммит.

После успешного `git push` прислать в чат краткий отчет: что сделано, какие проверки прошли, путь к zip и ссылка на commit.

## Запрет на изменение иконок Бэкап

Иконки undo/redo в блоке `Бэкап` утверждены как финальный фронтовый вид v32. Без прямого запроса пользователя и отдельного подтверждения не менять SVG-параметры в `style.css`.

Контрольный скриншот:

```text
screenshots/backup-left-undo-icon-approved-2026-08-28.png
```
