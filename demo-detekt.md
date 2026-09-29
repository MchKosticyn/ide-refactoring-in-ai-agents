# Демо-кейс: detekt, `compileAndLint` → `lint`

## Кейс

detekt — линтер Kotlin. `detekt-test` — тестовый DSL, которым авторы правил пишут спеки. Две extension-функции на `Rule` с одинаковыми параметрами:

- `compileAndLint(content, compilerResources)` — при `-Pcompile-test-snippets=true` сначала компилирует сниппет, потом линтит;
- `lint(content, compilerResources)` — только линтит (для сниппетов, которые не компилируются: `expect class`, обрывки без контекста).

PR https://github.com/detekt/detekt/pull/7873 (2025-02-20) слил их в одну: `Rule.lint(content, compilerResources, compile: Boolean = true)`, `compileAndLint` удалён. Чтобы старые вызовы `lint()` не начали компилировать, автор вручную проставил `compile = false` в 75 местах. Файл: `detekt-test/src/main/kotlin/io/gitlab/arturbosch/detekt/test/RuleExtensions.kt`.

Масштаб до PR: `compileAndLint` — 1428 вызовов, 131 файл; старый `lint(content)` — 78 вызовов в 17 файлах; не трогать: `lint(ktFile: KtFile, …)` (36 вызовов) и `FormattingRule.lint(content, fileName)` в `detekt-formatting` (66 вызовов, отдельная функция с тем же именем).

Ловушки:
1. Целевое имя занято с той же сигнатурой. Замена имени первой даёт две одинаковые `fun Rule.lint` → «Conflicting overloads»; после слияния в коде нет признака, какие из 78 вызовов были без компиляции. Тесты зелёные: обычный `gradle test` сниппеты не компилирует, ловит только CI-job с флагом.
2. Двойник `FormattingRule.lint` — какой `lint` вызовется из `subject.lint(code)` в `detekt-formatting`, решает строка импорта.
3. Вызовы без точки: `lint("fun f() { $code }")` внутри extension-функции (`BracesOnIfStatementsSpec.kt:2212`, `BracesOnWhenStatementsSpec.kt:1234`) — grep по `\.lint(` их не видит.

## Воспроизведение

```bash
git clone --filter=blob:none https://github.com/detekt/detekt.git && cd detekt && git checkout a4ec32a2a496439a2580951f4c340146a0e3d610
```

JDK 21, Gradle 8.12 (wrapper). IntelliJ: открыть корень, дождаться Gradle sync. Между режимами:

```bash
git checkout -- . && git clean -fdq
```

Промпт — описание PR словами автора (issue у PR нет), одинаковый для всех режимов:

> Implement the following change in this repository.
>
> The idea behind this change is that we don't want usages of the old `lint` because it doesn't checks that the code is "real" kotlin. This is bad because we assume that the code is correct to do our analysis so the test could be wrong. But we also know that there are cases where it's easier to just use `lint` with code that doesn't compile for multiple reasons.
>
> What to do is to rename `compileAndLint` to `lint` so the function is always the same AND if you in one case doesn't want to compile the snippet of code you can set `.lint(code, compile = false)` to specify that you don't want to verify the snippet. Existing tests must keep their current behavior. Leave `lintWithContext` / `compileAndLintWithContext` as they are.

Первые два абзаца — дословно из https://github.com/detekt/detekt/pull/7873; добавлена только рамка задачи. Ограничений на сборку нет: обычный `gradle test` ловушку не ловит (сниппеты компилируются только с `-Pcompile-test-snippets=true`), поэтому запуск Gradle агенту не помогает, а метрики становятся честнее.

Прогоны 2026-09-29 делались с другим, более подробным промптом (в `research/repro-2026-09-29.md`); цифры ниже — по нему.

Проверка результата:

```bash
git diff --stat | tail -1; git grep -c 'compile = false' -- '*.kt' | awk -F: '{s+=$2} END {print "compile=false:", s, "(эталон 75)"}'; git grep -n 'compileAndLint(' -- '*.kt' | grep -vc WithContext; git grep -nE '^\s+(val findings = )?lint\("fun f\(\)' -- '*.kt'
```

Ожидание: ~134 файла; `compile=false` ≈ 75; остатков `compileAndLint(` — 0; две строки `lint("fun f() { $code }"` должны содержать `compile = false`.

```bash
./gradlew :detekt-rules-empty:test :detekt-rules-style:compileTestKotlin :detekt-formatting:compileTestKotlin -Pcompile-test-snippets=true
```

Падение `EmptyDefaultConstructorSpec` при включённом флаге означает, что `compile = false` не проставлен.

## Что показали прогоны 2026-09-29

Claude Code, Opus, `--dangerously-skip-permissions`, MCP отключены: 13,7 мин, 53 хода (Bash 39, Read 4, Edit 9), $4.41, 133 файла. Ловушки 1 и 2 агент разобрал сам. Ловушка 3 — пропущены оба вызова без точки (искал `\.lint(`); последствие — эти хелперы стали компилировать сотни сниппетов под флагом (в коде комментарий «not compileAndLint for performance reasons»). Побочно: замена превратила этот комментарий в «not lint for performance reasons». Сборка проходит — промах молчаливый.

LSP (JetBrains Kotlin LSP через Serena): `references` на старом `lint(content)` — 99 ссылок, включая оба вызова без точки; перегрузка `lint(ktFile)` и `FormattingRule.lint` — отдельные символы. Rename `compileAndLint → lint` отклонён ошибкой `-32602` «Function 'lint' is already declared in package 'test'»: конфликт сервер видит, но в протоколе только «сделал / ошибка» — ни списка конфликтов, ни «продолжить».

IntelliJ, ожидаемый порядок: Change Signature на старом `lint` — добавить `compile: Boolean`, значение в вызовах `false` (78 мест подставляются движком); Rename `compileAndLint → lint` — диалог конфликта «уже объявлена»; слить тела вручную. Проверить до записи: как именно IDEA показывает конфликт и что делает при «Continue».

## Полезное для кадра

- Один экран для ловушки 1: `RuleExtensions.kt`, строки 21–43 (обе функции рядом).
- Ловушка 3 в кадре: `BracesOnIfStatementsSpec.kt:2209–2216` — `private fun BracesOnIfStatements.test(...)` с `lint("fun f() { $code }")`.
- Компактный модуль для показа результата: `detekt-rules-empty` (7 файлов, 44 `compileAndLint` + 5 старых `lint`, среди них `expect class` в `EmptyDefaultConstructorSpec.kt:70–100`).
- Эталонный дифф: `git diff a4ec32a2 073efe649 -- detekt-rules-empty`.
- Метрики для сравнения режимов: ходы, время, токены, стартовый контекст, число `compile = false`, два вызова без точки.
