# Salary Baranchik Payroll App

Легка фронтенд утиліта для локального розрахунку зарплат співробітників з експортом у XLSX.

## Що вміє
- Зберігає дані в localStorage (офлайн, без бекенда)
- Розрахунок оплат для різних типів ставок (fixed / погодинна / % офіціант / hostess)
- Імпорт списку співробітників та годин (вставкою тексту)
- Гнучка верстка та стилі у модульних CSS файлах
- Експорт у Excel:
  - Якщо доступний ExcelJS – максимально точне форматування
  - Fallback на SheetJS (спрощений стиль)

## Структура
```
index.html        – кореневий документ
js/               – модульний JS (state, utils, parse, pay, ui, xlsx, app)
css/              – розбиті стилі (variables, base, layout, components, table, modal, excel-editor, utilities)
package.json      – скрипти та dev залежності
```

Головна точка входу: `js/app.js`. Файл `js/main.js` був дублем і видалений.

## Запуск локально
1. Встановити залежності (тільки dev інструменти): `npm install`
2. Запустити простий сервер: `npm start`
3. Відкрити: http://localhost:5173

## Скрипти
- `npm run format` – відформатувати Prettier'ом
- `npm run format:check` – перевірка форматування
- `npm run lint` – ESLint перевірка
- `npm run lint:fix` – автофікс
- `npm start` – статичний сервер для перегляду

## ESLint / Стиль коду
Базується на `eslint:recommended` + кілька правил (eqeqeq, curly). Попередження для `console` (окрім warn/error). Допускаються змінні з префіксом `_` як невикористані аргументи.

## Експорт у Excel
Кнопка "Експорт в Excel" намагається використати ExcelJS. Якщо бібліотека не завантажилась – переходить на SheetJS.

## Зберігання даних
Все у localStorage, ключі: `payroll_employees_v1`, `payroll_meta_v1`, `payroll_excel_layout_v1`, `payroll_settings_v1`.

## Подальші ідеї (не реалізовано)
- Типізація через JSDoc / TypeScript
- Тести обчислень (computePays)
- Імпорт/експорт CSV
- Резервні копії у file download

## Ліцензія
UNLICENSED (використання всередині команди / приватно).
