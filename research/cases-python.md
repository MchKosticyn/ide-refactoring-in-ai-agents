# Поиск кейсов: Python / pytest (агент, 2026-09-28)

Проверено по диффам, тарболам на pre-PR коммитах, трекерам. Непроверенное помечено «гипотеза».

## Кейс 1. CPython — раскол `test.support` на подмодули (bpo-40275, 3.10)

Ссылки: https://github.com/python/cpython/issues/84456; коммиты 0d00b2a5 «Add os_helper» (2020-06-10), 19 коммитов «Use new test.support helper submodules in tests» (GH-20824…GH-21785), d94af3f7 «Remove test helpers aliases» (2020-08-08), фиксы промахов c6f282f3 (GH-21785, 2020-08-08) и 490c5426 (GH-21811, 2020-08-10).

`test.support` — хелперы тестового набора CPython. Вынесли в `os_helper`, `import_helper`, `threading_helper`, `socket_helper`, `warnings_helper`. Два месяца держали алиасы, затем удалили — посыпались buildbot'ы.

Масштаб (коммит 07d81128): имён, переезжающих в `os_helper`, ~1533 ссылки `support.<name>`; `support.unlink` — 313 ссылок в 70 файлах, рядом 165 `os.unlink`; 177 ссылок без вызова (`self.addCleanup(support.unlink, fn)`); 21 файл с голым `from test.support import unlink`. Компактный срез: `EnvironmentVarGuard` — 80 ссылок в 36 файлах, 4 стиля обращения (`support.` 52, голое 18, `test_support.` 6, `test.support.` 4), 4 файла вне `Lib/test` (`distutils/tests`, `ctypes/test`, `tkinter/test`).

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Одноимённые `support.unlink`/`rmtree` рядом с `os.unlink`/`shutil.rmtree`; голое `unlink` | Менять только «те», поиск даёт смесь | — (вручную по файлу) |
| Ссылка без вызова: `addCleanup(test.support.unlink, …)` | Регэксп по `name(` не находит; `test__osx_support` сломался на таких строках | c6f282f3, через сутки: pablogsal «PR21771 has broken a considerable amount of buildbots» |
| Разные стили импорта (`import test.support`, `from test import support as test_support`) | Массовые PR правили одну идиому, пропустили остальные | 490c5426: `test__osx_support`, `test_importlib/*/test_case_sensitivity`, `test_selectors` (`support.make_bad_fd`) |
| Тесты, не бегущие в обычном CI (macOS-only, регистрозависимая ФС, `-u largefile`) | Промахи жили 2–3 дня | vstinner 2020-08-10 |
| Тесты вне `Lib/test` | grep по `Lib/test` не видит; `ctypes/test/test_loading.py` правили отдельно | d94af3f7 |

Текст vs движок: PyCharm Move (F6) `EnvironmentVarGuard` в `os_helper` обновляет все 80 ссылок в 4 стилях и вне `Lib/test`. Честно: внешние проекты, импортировавшие `test.support.TESTFN`, движок не спасёт.
Демо: `git checkout 07d81128`, открыть `Lib`, Move `EnvironmentVarGuard` (или `make_bad_fd`); `python -m test test__osx_support`, `test_getopt` любым 3.x. Java-аналог: Move static method с изменением static import.

## Кейс 2. PyTorch — `torch.testing.assert_allclose` → `assert_close`

Ссылки: RFC https://github.com/pytorch/pytorch/issues/61844; миграция https://github.com/pytorch/pytorch/pull/61841 → коммит 99203580 (2021-08-19, 26 файлов); удаление https://github.com/pytorch/pytorch/pull/87974 (2022-11) → ревёрт 8c1c6759; удаление https://github.com/pytorch/pytorch/pull/164560 (2026-08) → ревёрты 9bb4e0c3, 36994097.

Контракт различается: `*` после `actual, expected` (keyword-only), `equal_nan` False (было True), `check_dtype=True` (было False), другие допуски, numpy не принимается.

Масштаб: 26 файлов, ~120 вызовов внутри; ~1000 внешних по RFC.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Keyword-only: позиционные `assert_allclose(a, b, 1e-3, 1e-3)` → `TypeError` | Ломается только при запуске конкретного теста | PR #61841 |
| Тихая смена дефолтов `equal_nan`, `check_dtype`, допусков | Тест меняет смысл молча | ревёрт 36994097: «Not a mechanical rename either: the removed wrapper had different default tolerances and passed equal_nan=True, check_dtype=False, check_stride=False, so callers need a semantics-preserving migration first» |
| Numpy/скаляры на входе | `torch.from_numpy(...)` правки | 99203580 |
| Сотни немигрированных вызовов | Удаление ревёртили трижды за 4 года | 2022-11, 2026-08 ×2 |

Движок: Change Signature — добавить параметры с явными дефолтами во все вызовы, добавить `*`, перевести позиционные в keyword; потом rename безопасен. Честно: numpy и допуски — человек.
Демо: статически да (checkout `99203580~1`), прогон — нет (сборка torch). Java-аналог: JUnit4→5 порядок аргументов `assertEquals`.

## Кейс 3. OpenStack Tempest — два одноимённых модуля `data_utils`

Ссылки: https://bugs.launchpad.net/tempest/+bug/1628016 (2016-09-27); фикс 02446973 (2017-02-14); консолидация 8 коммитов «Use tempest.lib data_utils — …» (2017-03-10, напр. 757833a2); удаление старого bd65bbb7; хвосты c0f9556c, 0f1e5cfe (2017-07-26, `test_oauth_consumers.py` чинили третий раз).

Tempest — фреймворк интеграционных тестов OpenStack, команда QA. Два модуля `tempest.common.utils.data_utils` (прокси, `rand_name` с `CONF.resources_prefix`) и `tempest.lib.common.utils.data_utils` (без префикса). Вызовы `data_utils.rand_name(...)` идентичны.

Масштаб (перед 757833a2): 154 файла импортируют старый, 20 — новый, 570 вызовов `rand_name`, ~270 других. Ни один файл не импортирует оба.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Одноимённый модуль в двух пакетах с разным поведением | Ресурсы без префикса не убирались cleanup'ом → утечки в облаке | LP#1628016; фикс через 4,5 мес |
| При обратной миграции один файл трижды менял импорт из-за конфликтов слияния | Старый импорт воскресал после мержа | 0f1e5cfe |
| Удаление прокси ломает плагины в других репо | гипотеза | 2017 |

Движок: Find Usages на символе из `tempest/common/utils/__init__.py` — 154 файла; Safe Delete прокси перечислит все. Честно: разницу поведения покажет человек.
Демо: очень просто, маленькое репо, unit-тесты бегут без облака. Java-аналог: два `TestUtils` в разных пакетах.

## Кейс 4. Home Assistant — фикстура `calls` → общая `service_calls`

Ссылки: https://github.com/home-assistant/core/pull/118349, /118350 (2024-05-29, 43+45 файлов); переименования чужих `calls`: #118353 calendar, #118354 mqtt, #118355 components, #118358 template (`service_calls` → `call_service_events`); общая фикстура + pylint-плагин #118356 (2024-06-08); 37 PR «Use service_calls fixture in …» (июнь–июль 2024); #120923.

~80–90 модулей имели копию фикстуры `calls`; в ряде модулей `calls` значило другое (mqtt — `MagicMock`, calendar — события).

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Одноимённые фикстуры с разной семантикой | Правка «всех `calls`» задела бы их | заранее: 4 PR-переименования |
| Целевое имя занято: `service_calls` в `tests/components/template/*` — фикстура другого типа | Локальная тихо перекрывает общую из `conftest.py` — тесты зелёные, проверяют другое | #118358 до введения общей |
| Строки: `@pytest.mark.parametrize(("fixture", …, "calls", …))` | argname перекрывает фикстуру | #118355 → `call_count` |
| `calls` — обычное слово (`mock.mock_calls`) | Замена по слову невозможна | — |

Движок: PyCharm резолвит фикстуры как символы, Rename обновляет `usefixtures`/`parametrize`, конфликт подсветит. Честно: консолидация 90 копий — серия шагов.
Демо: тяжело (репо огромное); один модуль — ок. Java-аналог: `@MethodSource("name")`.

## Отвергнуто

- Django `assertFormError` (ticket #33348, 50e1e7ef, 2021-12): смена сигнатуры, Change Signature не выведет `response.context[name]` — движок тоже проигрывает. Пример честного «и движок не спасёт».
- NiceGUI #3411 — откатили внутри PR.
- pytest-промахи в `usefixtures`/`getfixturevalue` — CI ловит до мержа, истории нет.

## Таблица (0–2)

| Кейс | Текст ломается / движок берёт | История промахов | Компактность | QA | Строки/мультиязычность |
|---|---|---|---|---|---|
| CPython `test.support` | 2 | 2 | 2 | 2 | 1 |
| PyTorch `assert_close` | 2 | 2 | 0–1 | 2 | 0 |
| Tempest `data_utils` | 1–2 | 2 | 2 | 2 | 0 |
| Home Assistant `calls` | 2 | 1 | 1 | 2 | 2 |
| Testcontainers | 1 | 1 | 2 | 1 | 0 |

Рекомендация агента: CPython основной (срез `EnvironmentVarGuard`/`unlink`), PyTorch — история про смысл (цитата ревёрта), Tempest — запасной.
