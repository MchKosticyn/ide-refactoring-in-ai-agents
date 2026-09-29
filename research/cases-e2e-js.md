# Поиск кейсов: Cypress/Playwright e2e-сьюты на JS/TS (подагент, 2026-09-28)

Проверено по `gh` API и sparse-клонам (metabase, gutenberg). Mattermost и matrix-react-sdk — только API, счётчики не сняты. Zulip — миграции на Playwright не было, кейса нет.

## Кейс 1 (лучший). Metabase — именованные импорты helpers → namespace `H.`

~200 глобальных helper-функций Cypress-тестов (`popover()`, `modal()`, `restore()`, `visitQuestion()`, `filter()`, `sidebar()`) переведены на один объект `H`; затем `cy.createQuestion` → `H.createQuestion`; затем `import { H }` → `const { H } = cy`.

- https://github.com/metabase/metabase/pull/50156 «Use global e2e helper import», 2024-12-04, 265 файлов, jscodeshift-кодмод (текст в PR).
- https://github.com/metabase/metabase/pull/52878 «Remove deprecated custom Cypress API commands», 2025-01-29, 158 файлов, perl-регекс `s/cy\s*\.\s*createQuestion/H.createQuestion/g` + ручной «fix `(cy as any)` block».
- https://github.com/metabase/metabase/pull/52042 `import { H }` → `const { H } = cy`, 2025-01-30, 284 файла.
- https://github.com/metabase/metabase/pull/61403 «Use H helpers instead of commands», 2025-07-27, 19 файлов — хвосты через 7 месяцев.
- https://github.com/metabase/metabase/pull/62662 «fix: remove helpers from the types of `cy`», 2025-09-01 — типы обещали `cy.activateToken`, которого нет в рантайме (7 месяцев).

Язык: JS + TS вперемешку, ~290 spec-файлов. Масштаб на `ad3fc72b037^`: `popover` 200 файлов / 2177; `filter` 140 / 2152; `restore` 262 / 1166; `modal` 125 / 794; `visitQuestion` 119 / 649; `sidebar` 114 / 498; `saveDashboard` 56 / 267.

| Ловушка | Что произошло | Когда поймано |
|---|---|---|
| Одно имя — helper и локальная функция: глобальный `filter({mode})` и локальные `function filter(label: string)` в `sql-filters-reset-clear.cy.spec.ts` (63 вызова) и `dashboard-filters-reset-clear.cy.spec.ts` (61); `sidebar()` vs `function sidebar()` в `embedding-smoketests.cy.spec.js` и `const sidebar = "sidebar open"` | Автор кодмода вручную внёс `filter`, `sidebar`, `visitQuestion`, `saveDashboard`, `duplicateTab` в `excludedValues` («problematic values to exclude»). Bare `filter(` до — 230, после — 125 (локальные), `H.filter(` — 105 | До мерджа автором; текстовая замена сломала бы 124 вызова молча |
| Имя helper'а = ключ опций: `visitQuestion()` (208) и `{ visitQuestion: true }` — 359 вхождений как ключ; `filter:` MBQL-ключ — 112; `saveDashboard:`/`duplicateTab:` — ключи событий Snowplow | Regex по слову даёт синтаксическую ошибку или пропуск | До мерджа |
| Три места кастомной команды: `Cypress.Commands.add("createQuestion", …)` + `interface Chainable` + `cy.createQuestion(` (296) при `H.createQuestion(` (339) и bare (12); `(cy as any).createQuestion` | Perl покрыл `cy.X`, `(cy as any)` руками, «pcregrep остатков» | До мерджа, ручной шаг |
| Дрейф типов: `interface Chainable extends HelperTypes` → IDE предлагала `cy.activateToken`, в рантайме `undefined` | Slack-жалоба, фикс #62662 через 7 мес | После мерджа |
| Хвосты кодмода: 9 файлов с прежними импортами (в т.ч. `import * as H from …` — namespace-импорт кодмод не матчил); к #61403 — 13 файлов (3 старых + 10 новых по старому образцу) | Повторили кодмод | После мерджа |

Демо (WebStorm): `sql-filters-reset-clear.cy.spec.ts` (локальная `function filter(label)`, 63 вызова) + файл с глобальным `filter()` и `{ visitQuestion: true }`. Replace in Path ломает, рефакторинг через resolve — нет. Java-аналог прямой: `import static Helpers.filter` при локальном `filter(String)` и поле `filter` в DTO.

## Кейс 2. Mattermost — `cy.apiAdminLogin()`: `.d.ts` разошёлся с реализацией, тест 2 года получал не то

Кастомные команды регистрируются строкой `Cypress.Commands.add('apiAdminLogin', fn)`, типы — вручную в `*.d.ts`.
- 2022-10-19 https://github.com/mattermost/mattermost-webapp/pull/11390: спек на TS, `cy.apiAdminLogin().then(({user}) => sysadmin = user)` — верно.
- 2022-11-03 https://github.com/mattermost/mattermost-webapp/pull/11479 «Fix apiadminlogin ts spec»: «исправил» `.d.ts` на `Chainable<UserProfile>` и спек на `.then((user) => …)` → `sysadmin.id` = `undefined`. tsc доволен.
- 2024-10-28 https://github.com/mattermost/mattermost/pull/28347 (41 файл): `user.js` → `user.ts`, типы `typeof apiAdminLogin` — компилятор увидел, спек вернули.

| Ловушка | Что произошло | Когда поймано |
|---|---|---|
| Три места: строка `Commands.add`, ручной `.d.ts`, вызовы | Тип «починен» в сторону ошибки, тест молча сломан | Через 2 года |
| `({user})` vs `(user)` | Diff в 4 символа меняет семантику | То же |

Гипотеза: тест `@enterprise @mfa @not_cloud` не гонялся в основном CI. Также: миграцию Cypress → Playwright в Mattermost делает Cursor-агент (#37454 2026-09-09, #37455 2026-09-19) — контекст «агенты уже мигрируют QA-код».

## Кейс 3. Gutenberg — одно имя `insertBlock` в двух пакетах с разной семантикой

`@wordpress/e2e-test-utils` (Puppeteer) и `@wordpress/e2e-test-utils-playwright` сосуществовали 2022–2024. На trunk@2023-06-01 (`74ea610fb9`): `insertBlock` — 109 (Puppeteer, `insertBlock('Site Title')` по UI-имени) vs 322 (`editor.insertBlock({ name: 'core/site-title' })` redux-dispatch по slug); `createNewPost` 96 vs 113; `pressKeyWithModifier` 180 → `pageUtils.pressKeys('primary+a')` (https://github.com/WordPress/gutenberg/pull/49009, 2023-03-14, 33 файла); `editor.canvas` `Frame | Page` → `FrameLocator` под тем же именем (https://github.com/WordPress/gutenberg/pull/54911, 2023-10-05, 85 файлов). Тесты — `.spec.js`, типов нет. Промахов после мерджа не найдено.

## Кейс 4 (слабый). Element / matrix-react-sdk — Cypress → Playwright, ~50 PR

Rewrite по файлам, не rename. Гипотеза: `createRoom` в трёх слоях с разными возвратами. Промахи — только «Deflake».

## Рекомендация агента

Сцена: Metabase `filter`/`sidebar`/`visitQuestion` (WebStorm). История №2: Mattermost `apiAdminLogin`. История №3: Gutenberg `insertBlock`.
