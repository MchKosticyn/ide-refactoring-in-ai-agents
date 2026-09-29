# Поиск кейсов: GitLab `qa/` (подагент, 2026-09-28)

Метод: blobless-клон `gitlab.com/gitlab-org/gitlab`, sparse `qa/`, worktree на pre-change коммитах, grep. Все три кейса проверены по коммитам и описаниям MR/issue.

## Кейс 1. `project_name_content` → `'project-name-content'` — пропущены EE-prepend-модуль и компонент-миксин

Миграция `data-qa-selector` (символы, snake_case) → `data-testid` (строки, kebab-case), осень 2023.
- MR https://gitlab.com/gitlab-org/gitlab/-/merge_requests/133010 (2023-10-05, `f43d489c9e56`): `app/views/projects/_home_panel.html.haml` + `qa/qa/page/project/show.rb`.
- Падение: https://gitlab.com/gitlab-org/gitlab/-/issues/427513 (2023-10-06, nightly, 72 отчёта) — `wait_for_import_success` → `WaitExceededError`.
- Фикс: https://gitlab.com/gitlab-org/gitlab/-/merge_requests/133550 (2023-10-10): «The selector used in this test has been recently updated in !133010… Since this test is fast-quarantined it did not fail in the original MR».

Масштаб (`f43d489c9e56^`): 5 ссылок в `qa/` + 2 в haml. Изменены 3 в `show.rb`; пропущены `qa/qa/ee/page/project/show.rb:30` (через `prepend_mod_with`) и `qa/qa/page/component/import/gitlab.rb:35`; `app/views/groups/projects.html.haml:31` — другой элемент с тем же именем, правильно не тронут.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Символ используется вне объявляющего класса: EE-модуль (`prepend_mod_with`), компонент | 2 из 5 ссылок остались; оба атрибута в DOM отсутствуют → 120 с и падение | Через сутки nightly; фикс через 5 дней |
| Одноимённый селектор на другой странице | Слепая замена задела бы (гипотеза) | — |
| Статическая проверка `sanity/selectors.rb` видит только `view … element` блоки | Прошла зелёной | — |

Честно: grep находит все 7 мест; автор проиграл «правлю только свою область». RubyMine найдёт 5 ruby-символов, haml-строку — нет. Кейс «строки — слабое место обоих». Для демо — средняя пригодность, для истории — хорошо.

## Кейс 2. Миграция auth-селекторов: MR → revert в тот же день → двухфазная миграция

- https://gitlab.com/gitlab-org/gitlab/-/merge_requests/135327 (2023-11-02, 57 файлов) → revert https://gitlab.com/gitlab-org/gitlab/-/merge_requests/135898 (тот же день): все QA-джобы staging-canary упали — `login-page did not appear on QA::Page::Main::Login` (`required: true`).
- Причина: QA-сьют гоняется против развёрнутого стенда со старым HTML; page object ищет `data-testid="login-page"`, страница отдаёт `data-qa-selector="login_page"`.
- Фикс: !136315 «Add testid to login page» (2023-11-08, testid рядом со старым), re-migrate !137104 (2023-11-21).

Для «текст vs движок» — низкая ценность (ловушка процессная), для сторителлинга — высокая.

## Кейс 3. `wait` → `wait_until` (+ `max:` → `max_duration:`, `interval:` → `sleep_interval:`)

- Подготовка `8802b61913eb` (2020-01-09); переименование https://gitlab.com/gitlab-org/gitlab/-/merge_requests/22861 (2020-01-15, `98124bd7f98e`, 57 файлов); пропуск https://gitlab.com/gitlab-org/gitlab/-/merge_requests/23477 (2020-01-22) — `Support::Waiter.wait` в `group_saml_sso_spec.rb`.

Масштаб (`98124bd7f98e^`): 5 определений `def wait` (`Page::Base`, `Support::Waiter`, `Support::Page::Logging` — prepend в `Page::Base` при debug, `Resource::Base`); 119 вызовов в 63 файлах; должны остаться 77 Capybara-kwarg `wait:` в 34 файлах (тот же коммит добавляет новые `wait:`); ~25 в комментариях/shell.

| Ловушка | Что произошло | Когда поймали |
|---|---|---|
| Одно имя в 4 классах + prepend-override, сигнатуры разные | Переименованы согласованно | в MR |
| Тот же токен `wait` как kwarg Capybara (77 мест) | Слепой `\bwait\b` испортит; regex `wait[ ({]` спасает — сильная модель справится (гипотеза) | — |
| Параллельная ветка (!21721, merged через 2 дня) добавила вызов удалённого метода | `NoMethodError` только в рантайме EE-спека | Через 7 дней, !23477 |

Честно: в самом rename-коммите пропусков нет; промах — merge skew в динамическом языке. Движок: Rename `Page::Base#wait` с иерархией → override в `Logging`; `Support::Waiter.wait`, `Resource::Base#wait` — отдельные символы.

## Бонус: интерполируемые имена элементов — слепая зона для всех

- `click_element(:"listbox_item_#{status}")` → фикс !123494 (2023-06-14).
- `click_element(:"listbox-item-#{target_group_path}")` → !128524 (2023-08-10), issue #420722.
- haml и vue версии одной кнопки с разными testid → !142454 (2024-01-22), через 6 дней после !141904.

## Итог агента

1. Кейс 3 — лучший по механике (одно имя в 4 классах, prepend, kwarg-коллизия), промах — merge skew.
2. Кейс 1 — лучшая история промаха, но движок не выигрывает.
3. Кейс 2 — драматично, нулевая ценность для «текст vs движок».
Не найдено: rename одноимённых методов в многих page objects с частичным изменением и задокументированным пропуском.

## Уточнения из финального отчёта того же агента
- Кейс 3 (`wait` → `wait_until`): kwargs `max:`/`interval:` переехали в `max_duration:`/`sleep_interval:`, причём `retry_until` уже использовал `max_duration:` — существующее имя-цель. После коммита bare-вызовов `wait` не осталось (leftover 0). Пропуск — только merge-race (верифицировано).
- Кейс 1: в том же MR CE-метод `wait_for_import` удалили («Remove unused wait_for_import test method»), EE-двойник `wait_for_import_success` не заметили. `Element#selector_css` матчит `[data-testid=NAME],[data-qa-selector=NAME]` — смена snake→kebab делает старый символ мёртвым без ошибки.
- Кейс 2 (auth): `element :login_field` → `'username-field'` — вью уже имела `testid: 'username-field'` рядом со старым (цель занята); `password_field` объявлен в трёх `view`-блоках.
- Финал миграции: после отключения fallback на `data-qa-selector` (!141904, 2024-01-16) нашлись остатки — !142445 (10 файлов), !142454 (haml и vue с разными testid).
- Кейс 3, ещё: `Support::Waiter.wait_until` уже существовал с другой сигнатурой (целевое имя занято); слепая замена `Support::Waiter.wait {}` → `wait_until {}` меняет семантику — дефолт `raise_on_failure` (коммит `c2a378d2547f` «Raise on failure by default», через 9 дней). Rename-коммит сам исправил 2 вызова `Support::Waiter.wait` в том же SAML-спеке (строки 338/347), пропущенный — из параллельной ветки.
