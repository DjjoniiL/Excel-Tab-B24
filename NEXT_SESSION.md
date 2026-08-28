# NEXT_SESSION

Обновлено: 2026-08-28

## Финальное Состояние

Проект `DjjoniiL/Excel-Tab-B24` доведен до `Excel Tab B24 MVP Final v.30`.

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
dist app B24 zip/Excel Tab B24 MVP Final v.30.zip
```

## Что Считать Актуальным

- Приложение остается serverless и browser-only.
- Runtime-файлы: `install.html`, `install.js`, `install.css`, `index.html`, `app.js`, `style.css`.
- Основное описание проекта: `README.md`.
- Подробная спецификация функций: `PROJECT_SPECIFICATION.md`.
- Правила интерфейса: `DESIGN_GUIDE.md`.
- Верхнеуровневый учебник обновлен рекомендациями из v26-v30.
- В Marketplace zip не входят документация, тесты, `.git`, `.env`, `node_modules`, локальные данные и служебные файлы.

## Права Приложения

- CRM (CRM)
- Placement / Встраивание приложений
- `user_basic`

Полный `user` и `user.userfield` не нужны.

## Финальный Контроль

Перед любыми будущими изменениями сначала читать:

- `README.md`
- `PROJECT_SPECIFICATION.md`
- `DESIGN_GUIDE.md`
- верхнеуровневый `Bitrix24 CRM Marketplace App Logic.md`

В этом handoff-файле нет открытых задач на следующую сессию. Новые задачи должны появляться только из нового запроса пользователя.

## Запрет На Изменение Левой Иконки Бэкап

Левая иконка undo в блоке `Бэкап` утверждена 2026-08-28. В следующих сессиях не трогать её дизайн без прямой просьбы пользователя. Утверждённое состояние:

- реализация: `style.css`, селектор `.icon-button.undo-button`;
- контрольный скриншот: `screenshots/backup-left-undo-icon-approved-2026-08-28.png`;
- указатель расположен на нижнем конце дуги и смотрит вправо;
- дуга тонкая, визуально опущена ниже и упирается в заднюю часть указателя.
