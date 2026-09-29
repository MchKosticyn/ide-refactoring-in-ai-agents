# Поиск кейсов: UI-автотесты и page objects (агент, 2026-09-28)

Проверено по фактическим диффам и снимкам файлов на pre-PR коммитах. Не проверены: GitLab `qa/`, Keycloak/Jenkins ATH, Cypress→Playwright миграции.

## Кейс 1. Firefox for Android (Fenix), Kotlin — шесть одноимённых `verifySnackBarText`

Проект: `mozilla-mobile/firefox-android`, `fenix/app/src/androidTest/.../ui/robots/**` (Robot = page object, QA-owned).
PR: https://github.com/mozilla-mobile/firefox-android/pull/5371 (merged 2024-02-01, +13/−32, 15 файлов). Проблема замечена в https://github.com/mozilla-mobile/firefox-android/pull/5313 (2024-01-30).

Шесть методов `verifySnackBarText(expectedText: String)`: `TestHelper` (object), `BrowserRobot`, `CollectionRobot`, `HomeScreenRobot`, `SettingsSubMenuHomepageRobot`, `TabDrawerRobot`. Пять делают assert, `CollectionRobot` — только `waitForExists`, без проверки. PR удаляет пять версий, оставляет `TestHelper`, добавляет импорты.

Масштаб (base `211c245e`): 72 вызова в 15 файлах + 6 определений. Вызовы текстуально идентичны; какой метод вызывается — определяется ресивером DSL-лямбды (`collectionRobot {}` vs `browserScreen {}`).

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Одноимённые методы в 6 классах, разрешение по ресиверу лямбды | Из-за «пустой» версии два теста (`CollectionTest`, `ComposeCollectionTest`) проверяли неверный текст `"Tabs saved!"` вместо `"Collection saved!"` и годами были зелёными | #5313 → #5371 |
| Удаление метода при наличии top-level одноимённого | Вызов тихо переезжает на импортированный `TestHelper.verifySnackBarText`, если импорт есть; иначе compile error только на `assembleAndroidTest` | В PR — импорты в 9 файлов |

Текст: grep даёт 78 одинаковых строк, не говорит, какая версия вызывается. Движок: Find Usages / Safe Delete на `CollectionRobot.verifySnackBarText` показывает ровно его вызовы; Inline делает видимым отсутствие assert. Честно: ошибочные строки движок не найдёт.
Демо: Android-монорепо, нужен SDK и Gradle-sync, репо архивировано (код в `mozilla-central/mobile/android/fenix`). Как слайд — отлично, как демо завтра — рискованно.

## Кейс 2. Uyuni (SUSE Manager), Ruby + Cucumber — два промаха при переименовании шага

Проект: `uyuni-project/uyuni`, `testsuite/` (Ruby/Capybara, QA-owned).
Цепочка: https://github.com/uyuni-project/uyuni/pull/3227 (2021-02-18) → https://github.com/uyuni-project/uyuni/pull/3300 (2021-02-19, «restores the step … which was destroyed by accident») → https://github.com/uyuni-project/uyuni/pull/3333 (2021-03-02, «Fix: Wrong renaming of a step»).

В #3227 автор отредактировал на месте `When(/^I enter "([^"]*)" as the filtered system name$/)` → `… filtered package states name`. Старый шаг вызывался из другого step definition через `steps %( When I enter "#{…}" as the filtered system name )` (`navigation_steps.rb:830`) — grep не нашёл из-за интерполяции. Undefined step в рантайме → восстановлен в #3300. В #3300 новое определение получило `package states name`, feature-файлы вызывали `package name` → #3333 через 11 дней.

Масштаб (base `7c62960d`): `filtered package name` — 24 использования в 12 файлах; `filtered system name` — 1 определение + 1 вложенный вызов.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Строковая связь Gherkin ↔ regex step definition | Правка regex на месте уничтожила используемый шаг | На следующий день (#3300) |
| Вложенный вызов через интерполированную строку `steps %(…)` | Grep по полному тексту не находит | #3300 |
| Два похожих шага (`package name` / `package states name`) | Новое определение с не тем суффиксом | Через 11 дней (#3333) |

Движок: Cucumber-плагин JetBrains резолвит строки шагов на step definitions — Rename переписывает feature-строки, неразрешённый шаг подсвечен «Undefined step». Честно: интерполированный `steps %()` не резолвит ни grep, ни IDE — только `cucumber --dry-run`.
Демо: `testsuite/` открывается в RubyMine без сборки продукта. Java-аналог: cucumber-jvm `@When("…")`, IntelliJ резолвит так же.

## Кейс 3 (история, частично). Grafana — `@grafana/e2e-selectors`

https://github.com/grafana/grafana/pull/129516 (2026-07-30, 117 файлов): rename edit pane → sidebar; ключи селекторов — публичный API трёх кодовых баз (app, e2e, внешние плагины), поэтому старые оставлены `@deprecated`, новые версионированы. Серия «Migrate E2E to page objects» #1–#23 и cleanup #130211/#130221/#130497 (август 2026). Истории промахов не найдено. Ловушка «уникальное имя» — та, что модель решает текстом.
Побочно (гипотеза): Kibana https://github.com/elastic/kibana/pull/175711 — rename `data-test-subj` + page objects, 64 файла.

## Таблица (0–2)

| Критерий | Fenix | Uyuni | Grafana | Testcontainers |
|---|---|---|---|---|
| Текст проваливается, движок справляется | 2 | 2 | 1 | 0–1 |
| История промахов | 2 | 2 | 0 | 1 |
| Компактность / воспроизводимость | 1 | 2 | 1 | 2 |
| Релевантность QA | 2 | 2 | 2 | 0 |
| Мультиязычность / строки | 0 | 2 | 2 | 0 |
| Сумма | 7 | 10 | 6 | 3–4 |

Рекомендация агента: Uyuni основной, Fenix — слайд-история, Grafana — вспомогательный слайд.

## Дополнение: Java-кейсы

### Кейс 0. Keycloak — `isCurrent()` → `assertCurrent()` при уже существующем `assertCurrent()`

Проект: `keycloak/keycloak`, page objects `testsuite/integration-arquillian/tests/base/src/main/java/org/keycloak/testsuite/pages/**` (Selenium/Arquillian, QA-owned).
PR: https://github.com/keycloak/keycloak/pull/49814 «Unify page objects between arquillian and test-framework», merged 2026-07-20, +1764/−2331, 252 файла, pre-PR коммит `471f2762`. Java, IntelliJ (Maven-модуль).

Подтверждено по коду на base-коммите:
- Целевое имя занято: `AbstractPage.assertCurrent()` (стр. 37) и `LoginPage.assertCurrent(String realm)` (249) существовали до PR.
- Перегрузка: `isCurrent(String expectedTitle)` в `AbstractPage` (45), `isCurrent(String realm)` в `LoginPage` (245).
- Иерархия: абстрактный `isCurrent()` переопределён в ~55 page-классах двух иерархий (цифра субагента, проверено 3 класса).
- Смена типа `boolean → void`: `assertFalse(page.isCurrent())` (~33 места) при механической замене → `page.assertCurrent()` — инверсия смысла теста.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| `assertFalse(x.isCurrent())` → `x.assertCurrent()` | Инверсия в `WebAuthnOtherSettingsTest`, `AttestationConveyanceRegisterTest`, `UserVerificationRegisterTest`, `OIDCAdvancedRequestParamsTest` | Copilot-ревью в PR (комментарии проверены через API) |
| Целевое имя занято | Слияние двух семантик вручную | автор |
| Перегрузка `isCurrent(String)` | Отдельно | в PR |

Движок: Rename → «method `assertCurrent()` already exists»; Change Signature `boolean → void` подсвечивает места, где результат используется как выражение.
Демо: ≤3 мин на `LoginPage` + `AbstractPage` + два теста; сборка не нужна, индексация модуля.

### Кейс 0b. Jenkins acceptance-test-harness (гипотеза, диффы не просмотрены)

- Коммит `35a883c3` (2017-03-31): `edit(Runnable)` → `configure(Runnable)` при трёх существующих перегрузках `configure(...)` и постороннем `private edit(...)` в `AbstractAnalysisTest`; хвост через 2 дня в `ed49e729`.
- https://github.com/jenkinsci/acceptance-test-harness/pull/1119: `withTimeout(long, TimeUnit)` → `Duration`, 39 мест / 25 файлов, 2 пропущены, дочищены в #1124.

### Обновлённая таблица (0–2)

| Критерий | Keycloak | Uyuni | Fenix | Jenkins ATH | Grafana | Testcontainers |
|---|---|---|---|---|---|---|
| Текст проваливается, движок справляется | 2 | 2 | 2 | 2 | 1 | 0–1 |
| История промахов | 2 | 2 | 2 | 2 (гип.) | 0 | 1 |
| Демо завтра | 2 | 2 (RubyMine) | 1 | 2 (гип.) | 1 | 2 |
| QA | 2 | 2 | 2 | 2 | 2 | 0 |
| Строки / мультиязычность | 0 | 2 | 0 | 0 | 2 | 0 |
| Сумма | 8 | 10 | 7 | 8 | 6 | 3–4 |

Рекомендация агента: Keycloak для живого демо в IntelliJ; Uyuni — слайд про строки; Fenix — резерв; Jenkins ATH — запасное Java-демо после проверки диффов.

## Дополнение 2: детальная проверка Java page-object сьютов (подагент)

### Keycloak #49814 — уточнения (pre-PR `471f27628430c2cd04cd8ef6f363592e3f3e8a41`)

Определений `isCurrent(` в `src/main` — 55; вызовов `.isCurrent(` в тестах — 270 строк в 64 файлах: `assertFalse(...isCurrent())` — 33, `if (...isCurrent(...))` — 13, с аргументом — 16. `.assertCurrent(` уже был — 1408 вызовов.

Дополнительная ловушка: две несвязанные иерархии с одним именем — `pages/AbstractPage` (~48 наследников) и `page/AbstractPage` → `auth/page/login/*` (7 определений). Вторую менять не надо — после PR там осталось 7 определений и 5 вызовов (`DeleteAccountActionTest`). Grep даёт 55, править ~48 — отличить только по типу.
Перегрузка: `assertCurrent(String)` в `RegisterPage` принимает `orgName` — другая семантика; `passwordPage.isCurrent("consumer")` при замене даёт несуществующий метод.
Copilot-ревью: 100+ комментариев, инверсии `OIDCAdvancedRequestParamsTest:910`, `AttestationConveyanceRegisterTest:59`, `UserVerificationRegisterTest:113`. Для нового test-framework — серия фиксов `assertCurrent()` матчит старую страницу с тем же page-id: #52077, #52311, #52316, #52416, issue #52074. Пропущенных вызовов в компилируемом коде нет (компилятор).
Демо: Maven-модуль `integration-arquillian-tests-base`; гипотеза — без `999.0.0-SNAPSHOT` зависимости красные, но Rename/Find Usages по исходникам работают. Подмножество для 3 мин: `LoginPage`/`PasswordPage`/`LoginUsernameOnlyPage` + `BrowserFlowTest` (13 `assertFalse`).

### Jenkins ATH — `edit(Runnable)` → `configure(Runnable)` (подтверждено)

https://github.com/jenkinsci/acceptance-test-harness/commit/35a883c3d4af7006a7831a0f4a62874d5f9981a9 (2017-03-31, PR #294), pre-commit `2f04255c6c8885a60de2414774fce94d9b8b425a`. `po/ConfigurablePageObject.java`; 6 вызовов в 4 тестах; 5 файлов.
- Целевое имя занято: `configure()`, `configure(Closure)`, `<T> configure(Callable<T>)` (строки 54/60/76) + 22 других `configure(` в `po/`. Лямбда с return уходит в `Callable`-перегрузку (оборачивает исключения в `AssertionError`) — тихая смена контракта.
- Чужой `private edit(...)` с 5 параметрами в `AbstractAnalysisTest` — 5 вхождений.
- Промах: через 2 дня `jenkins.edit(`/`security.edit(` в `WarningsPluginTest` — https://github.com/jenkinsci/acceptance-test-harness/commit/ed49e729e3bf8572fb765ae41ec35b49599fe763 (2017-04-01).
Демо: всё на одном экране, репо standalone.

### Jenkins ATH — `withTimeout(long, TimeUnit)` → `withTimeout(Duration)` (подтверждено)

https://github.com/jenkinsci/acceptance-test-harness/pull/1119 (2023-04-25), pre-PR `ccd9f0a22594db3f81def6d82ca793c81f878ef0`, 25 файлов. 44 вхождения: 30 `SECONDS`, 3 `MINUTES`, 2 `MILLISECONDS` с выражением, 1 переменная, 5 уже `Duration`. `Wait.java` делегирующую перегрузку менять нельзя; Selenium `FluentWait.withTimeout(Duration)` — одноимённый чужой. Промах: 2 вызова в `CapybaraPortingLayerImpl.java` (157, 184) → #1124 (2023-04-28). Движок: Inline Method на deprecated перегрузке.

### Fenix #5371 — уточнения (pre-PR `211c245e045cac6f31a38566b2de51f21809c86f`)

~66 вызовов в 15 файлах; оценка по receiver: BrowserRobot ≈25, TabDrawerRobot 12, HomeScreenRobot ≈12, CollectionRobot 4, SettingsSubMenuHomepageRobot 1, TestHelper ≈18. Receiver лямбды объявлен в другом файле (`ThreeDotMenuMainRobot.kt:386 fun addToFirefoxHome(interact: BrowserRobot.() -> Unit)`). Ложный ориентир: `ComposeTopSitesTest.kt` импортирует `TestHelper.verifySnackBarText`, но вызовы внутри `addToFirefoxHome {}` идут в член `BrowserRobot` (implicit receiver побеждает импорт). Риск: нужен Android SDK + Gradle sync для резолва.

### Keycloak #47878 (запасной) — удаление `LoginPage.open()`

https://github.com/keycloak/keycloak/pull/47878 (2026-04-09), pre-PR `5183e49a8f9bce20a690c6f80b438cf9d4f6c80f`, 81 файл. `loginPage.open()` — 472 вызова в 76 файлах → `oauth.openLoginForm()`; перегрузка `open(String realm)` — 253, трогать нельзя; наследник `loginUsernameOnlyPage.open()` — 33, grep по `loginPage.open` не видит; `.open()` всего 794 (263 fluent `AbstractUrlBuilder.open()`, 8 `appPage.open()`). Старый `open()` включал `assertCurrent()` — в PR проверка тихо исчезла. Промахов нет (компилятор).

### Не подошло
- keycloak #48240 `oauth.clientId()` → `client()`: новый метод обнуляет `clientSecret` — семантика, контрпример-слайд.
- ATH: массовых rename `find`/`clickButton` нет.
- WordPress-Android e2e, openmrs-contrib-qaframework — мелко.
- Fenix `480cb9800f` «Break out TestHelper» — грепается, слабая ловушка.
