# Поиск кейсов: Java/Kotlin тестовая инфраструктура (агент, 2026-09-28)

Всё сверено с диффами на pre-PR коммитах; непроверенное помечено «гипотеза».

## 1. detekt: `compileAndLint` → `lint`, когда `lint` уже существует

Проект: detekt (Kotlin-линтер), 100 % Kotlin. PR https://github.com/detekt/detekt/pull/7873 «Simplify test api», коммит `073efe649`, 2025-02-20 (Brais Gabín); продолжение https://github.com/detekt/detekt/pull/7966 (`cf68bc0eb`, 2025-02-21). Pre-PR sha: `a4ec32a2a496439a2580951f4c340146a0e3d610`.

`detekt-test` — тестовый DSL авторов правил: `rule.compileAndLint(code)` (при `-Pcompile-test-snippets=true` компилирует сниппет, потом линтит) и `rule.lint(code)` (только линтит). Обе — extension на `Rule` с одинаковыми параметрами. PR переименовал `compileAndLint` → `lint` с параметром `compile: Boolean = true`, старый `lint` удалил; чтобы старые `lint()` сохранили поведение, автор вручную дописал `compile = false` в 75 местах.

Масштаб (`a4ec32a`): `compileAndLint(` — 1429 вызовов в 131 файле (13 модулей); старый `lint(` — 178 в 38 файлах (`style` 67, `formatting` 63, `naming` 38, `empty` 5); PR — 134 файла, +1693/−1726. Компактный срез: `detekt-rules-empty` (7 файлов, 44 `compileAndLint` + 5 `lint`) + `detekt-formatting`.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Целевое имя занято: `fun Rule.lint(content, compilerResources)` с той же сигнатурой. После переименования 178 старых `lint()` неотличимы от новых и по умолчанию компилируют | 75 мест — `compile = false`, ~100 молча сменили семантику | не сломалось: `./gradlew test` сниппеты не компилирует; видит только CI-job `compile-test-snippets` |
| Двойник в другом модуле: `fun FormattingRule.lint(content, fileName)` в `detekt-formatting/…/TestFiles.kt` (63 вызова). Какой `lint` вызовет `subject.lint(code)`, решает строка импорта | 6 файлов переведены с импортом, 13 остались на двойнике | автором |
| Перегрузка, которую трогать нельзя: `fun Rule.lint(ktFile: KtFile, …)` | не тронута | — |
| Строка-сообщение: замена превратила `"use lint instead of lintWithContext"` в `"use lint instead of lint"` | попало в мастер | #7966 на следующий день, попутно |

Текст: `sed compileAndLint→lint` даёт два одинаковых `fun Rule.lint` → «Conflicting overloads»; агент сливает — и признака, какие из 178 были «без компиляции», больше нет. Тесты зелёные. Движок: Rename показывает «function `lint` already declared»; Find Usages старого `lint` даёт ровно 178 (двойник — другой символ). Честно: какие нуждаются в `compile = false` — решает человек.
Воспроизводимость: Gradle 8.12, JDK 21, чистый Kotlin — открывается штатно.

## 2. Micronaut Test: три аннотации `@MicronautTest` с одним именем

Java + Kotlin + Groovy. PR https://github.com/micronaut-projects/micronaut-test/pull/229 (`95d6799b`, 2020-09-23); баг-фикс https://github.com/micronaut-projects/micronaut-test/pull/303 (`11f593cf`, 2020-10-01); миграция тестов https://github.com/micronaut-projects/micronaut-test/pull/320 (`aee4972e`, 2020-10-27). Pre-PR: `d67244bd38e090bb0b2e3b1a923f703cad9a1e47`.

Общая аннотация разделена на двойников с тем же простым именем в `junit5`/`spock`/`kotest`; старая deprecated. 84 файла тестов (36 Java, 40 Kotlin, 17 Groovy) не мигрировали в #229.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Одно имя — разные классы: в `MicronautSpockExtension.java` остался импорт старой, код `spec.getAnnotation(MicronautTest.class)` под новую. Компилируется; спека получает `null` | баг в релизе v2.1.0 | через 8 дней, #303 — фикс в одну строку импорта |
| У `test-kotlintest` (17 файлов) не было своей аннотации | в #320 направили на junit5-двойника | удалением модуля в Micronaut 3 (#437, 2021) |
| Старая deprecated → компилятор не подсказывает | 84 файла жили на старой месяц | #320 |

Движок: split, не rename; резолв `MicronautTest.class` PSI видит, grep нет. Воспроизводимость: Gradle 6.6.1 → JDK ≤ 14 — риск.

## 3. REST Assured: JUnit 4 → 5 в два захода, `@Rule` молча перестал работать

Java + Groovy + Scala + Kotlin. Коммиты в master (Johan Haleby): `b6cd8419` 2025-11-14 (258 файлов), `3b310d76` 2025-11-21 (43 файла), `24380dbf` 2025-12-12 (откат osgi на JUnit 4). Pre: `2521f9d0287e7580e7680c232222c04f6e23b6e1` или `b6cd8419` (полумигрированное состояние).

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| `@Rule` в Jupiter-классе игнорируется: 6 файлов `modules/spring-web-test-client` с Jupiter-`@Test` и `@Rule ExpectedException`; `exception.expect(...)` ничего не делает | пропущено | тестами, через неделю |
| `ScalaITest.scala` — `@BeforeEach` от Jupiter, `@Test` от JUnit 4 | пропущено | через неделю; `Scala3ITest`, `KotlinITest.kt` тоже |
| OSGi мигрировать не удалось | — | откат через месяц |

Движок переименования не помогает; помогает инспекция JUnit-миграции. Кейс «история + QA», не «engine wins».

## 4. WireMock: `Request.getPathAndQuery` → `getPathAndQueryWithoutPrefix` при `Url.getPathAndQuery()`

https://github.com/wiremock/wiremock/pull/3288 (`234ddb1b`, 2026-01-20). Pre: `7b096a3ca257e6b3955ee1e3ccea43551eb1d9f2`. 39 упоминаний в 27 файлах; не менять: `Url.getPathAndQuery()`, `Origin.getPathAndQuery()`, `ImmutableRequest:65`, `OriginTests:52`, комментарий у deprecated `getUrl()`. Тест-дубль `MockRequest` и Mockito-стаб переименованы. Провал текста шумный (компиляция), пропусков нет. Gradle, JDK 17.

## 5. Kafka: `TestUtils.waitForCondition` — перестановка двух `long` в одной из пяти перегрузок

https://github.com/apache/kafka/pull/10759 (`e4b3a3cd`, 2021-05-27). Pre: `a02b19cb77084a573a25be1b75fc195b9a9c1f9b`. 279 вызовов, затронуто 5. Пропущенный вызов компилируется и ждёт 100 мс с опросом раз в 60 с. Три класса `TestUtils` (Java clients, Scala core, `test-common`) — реальный баг в третьей копии KAFKA-18214 (2024). Change Signature переставляет аргументы этой перегрузки. Kafka завтра не открыть.

Бонус в Kafka: EasyMock → Mockito в `streams` (https://github.com/apache/kafka/pull/10850, 2021-06-11): голые вызовы `addValueMetricToSensor(...)` из «записанных ожиданий» стали обычными вызовами — тесты зелёные, ничего не проверяют. Пять фиксов через год (#12322, #12323, #12325, #12373, #12454): «This miss happened during the switch from EasyMock to Mockito». Гипотеза: в `ClientMetricsTest:160`, `RocksDBMetricsTest` мёртвые ожидания до сих пор.

## Попутно (для слайдов «история пропусков»)

- Selenium `ae62ba00` (#10778, 2022-06, JUnit 4 → 5, 503 файла): собственная `org.openqa.selenium.testing.Ignore` (46 файлов) vs `org.junit.Ignore` (25); два `@Ignore` в `PerSessionLogHandlerUnitTest` без импорта — файл не в Bazel-target, заметили через 4 дня (#10793).
- Spring Framework Jupiter (`3877087f`, 2019-08, 1933 файла): переопределённые тест-методы без `@Test` Jupiter не запускает; follow-up'ы `aa6e762d`, `3e2b977d`; `@Inject`-TCK выпали на месяц (`4dc966b0`).
- Hibernate ORM `6e47b328` (HHH-20177, 2026-02): `@Rule LoggerInspectionRule` в Jupiter-тесте два года вхолостую.
- Keycloak `f2839b12` (#48493, 2026-04): тест перенесли между модулями, `@SelectClasses`-сьюты не обновили — три дня не гонялся.
- Mockito `5e052cbb` «Fixes forgotten renames» (2013-12): строки `"org.mockitoutil.…Test$ClassUsingInterface1"` для `loadClass` — Rename с «search in strings» находит.
- JUnit 5 `39e9e630` (2016-10): `expectThrows` → `assertThrows` при существующем `assertThrows` (void); регрессия в generic-сигнатуре, фикс `bf30a7bb` (#599). Шумно.
- cucumber-jvm: JUnit 4 → 5 в пять заходов (2019-08…2020-04), Kotlin-модуль пропущен дважды (`0e05ca42`); `TypeRegistry` — два класса с одним именем (`e12334cf`).
- Пусто: ktor, Exposed, Kotest, AssertJ, Allure, Playwright-java, Commons, Hibernate `@TestForIssue → @JiraKey`.

## Таблица (0–2)

| Кейс | Текст проваливается / движок ловит | История | Компактность / завтра | QA | Строки / языки | Σ |
|---|---|---|---|---|---|---|
| Testcontainers | 2 | 2 | 2 | 1 | 2 | 9 |
| detekt | 2 (молча, конфликт имени) | 1 | 1 | 2 | 1 | 7 |
| Micronaut Test | 1 | 2 | 1 (JDK ≤ 14) | 1 | 2 | 7 |
| REST Assured | 1 | 2 | 1 | 2 | 2 | 8 |
| WireMock | 1 | 0 | 2 | 1 | 0 | 4 |
| Kafka | 2 | 0 | 0 | 2 | 1 | 5 |

Рекомендация агента: detekt основной (Rename → конфликт → Find Usages 178 → `detekt-formatting` с двойником); REST Assured — слайд «текст + история»; Micronaut — иллюстрация; Kafka — лучший Change Signature, не воспроизвести.
