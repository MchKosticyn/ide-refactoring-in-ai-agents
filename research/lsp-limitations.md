# Что LSP даёт и не даёт для рефакторинга, когда клиент — AI-агент

Дата исследования: 2026-09-29. Источники: спецификация LSP 3.17 (стабильная на момент большинства
реализаций), 3.18 (выпущена 04.06.2026, «current»), черновик 3.19 (пустой changelog), issues репозитория
`microsoft/language-server-protocol`, исходники VS Code / vscode-languageclient и серверов — последние
только как иллюстрации.

Обозначения ссылок: `SPEC317 = https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/`,
`SPEC318 = https://microsoft.github.io/language-server-protocol/specifications/lsp/3.18/specification/`,
`LSPI = https://github.com/microsoft/language-server-protocol/issues/`.

Замечание по версиям: раздел `rename`/`prepareRename` в 3.18 отличается от 3.17 только пунктуацией и
введением именованных типов `PrepareRenameResult` и т.п.; семантика не изменилась (сверено diff'ом
текста разделов). В черновике 3.19 на сегодня ни одного `@since 3.19`.

---

## 1. `textDocument/rename` и `textDocument/prepareRename`

**Тезис 1.1. Результат rename — только `WorkspaceEdit | null`. Никакого поля для предупреждений, конфликтов или «список проблем, но вот правки».**
— «result: `WorkspaceEdit | null` describing the modification to the workspace. null should be treated
the same was as WorkspaceEdit with no changes (no change was required).»
— `SPEC317#textDocument_rename`

**Тезис 1.2. Единственный канал «отказать» — JSON-RPC error с кодом и строкой message. Это отказ целиком, а не «предупреждение + продолжить».**
— «error: code and message set in case when rename could not be performed for any reason. Examples
include: there is nothing at given position to rename (like a space), given symbol does not support
renaming by the server or the code is invalid (e.g. does not compile).»
— `SPEC317#textDocument_rename`
— Про невалидное имя: «The new name of the symbol. If the given name is not valid the request must
return a [ResponseError] with an appropriate message set.» (`RenameParams.newName`, там же).
Что такое «not valid» — не определено; синтаксически невалидное имя и имя, создающее конфликт, для
протокола неразличимы: оба — просто error с текстом.

**Тезис 1.3. `prepareRename` отвечает только «где» и «что», но не «почему нельзя» структурно. Причина отказа — только текст message в error.**
— «result: `Range | { range: Range, placeholder: string } | { defaultBehavior: boolean } | null`
describing a Range of the string to rename and optionally a placeholder text … If null is returned then
it is deemed that a 'textDocument/rename' request is not valid at the given position.»
— «error: code and message set in case the element can't be renamed. Clients should show the information
in their user interface.»
— `SPEC317#textDocument_prepareRename`
— Важно для агента: `prepareRename` вызывается ДО ввода нового имени (в параметрах только позиция:
`PrepareRenameParams extends TextDocumentPositionParams, WorkDoneProgressParams {}`), поэтому он
принципиально не может проверять конфликт с новым именем. Это подтверждает Dirk Bäumer:
«There is now also a prepareRename request. This is send before the user is actually asked for a new
name.» — `LSPI566`, dbaeumer, 2018-09-12.

**Тезис 1.4. Единственный протокольный механизм «предупредить, но дать продолжить» — `ChangeAnnotation.needsConfirmation` внутри `WorkspaceEdit` (с 3.16). Он привязан к правкам, а не к операции, и рассчитан на UI.**
— «A flag which indicates that user confirmation is needed before applying the change.»
(`ChangeAnnotation.needsConfirmation`); «A human-readable string describing the actual change. The
string is rendered prominent in the user interface.» (`label`).
— `SPEC317#changeAnnotation`
— Клиент должен заявить `workspace.workspaceEdit.changeAnnotationSupport`, а для rename — ещё и
`textDocument.rename.honorsChangeAnnotations`: «Whether the client honors the change annotations in text
edits and resource operations returned via the rename request's workspace edit by for example presenting
the workspace edit in the user interface and asking for confirmation.» — `SPEC317#renameClientCapabilities`.
— Если клиент (агент) этих capability не заявил — «If a client doesn't signal the capability, servers
shouldn't send AnnotatedTextEdit literals back to the client.» — `SPEC317#textDocumentEdit`. То есть
агент, не заявивший поддержку, предупреждений не получит вовсе.
— Никакой машиночитаемой категории конфликта (kind/code/severity) в `ChangeAnnotation` нет: только
`label`, `description`, `needsConfirmation`.

**Тезис 1.5. Мейнтейнеры сознательно не добавляли в rename поток «подтверждение/продолжить всё равно»; позиция — многошаговость только через данные, без UI из сервера.**
— Проблема названа ещё в 2016: «communication on sematic shift warnings. Refactoring should be semantic
preserving. But sometimes the semantic shift is wanted and we need to let the user confirm this.» —
`LSPI61`, dbaeumer, 2016-10-31.
— Кейс «имя занято → спросить "Rename anyway"» обсуждался в `LSPI1104` (DanTup, 2020-10-14):
сервер спросил через `showMessageRequest`, пользователь отменил — протокол не определяет, что вернуть.
Ответ dbaeumer 2020-10-20: «I think the right approach is to use `RequestCancelled` (or if we think it
is necessary a new `RequestUserCancelled`) in the response.» Ответ jrieken (VS Code) 2020-10-20:
«I think the problem begins with showing a "custom" message with options. You leaving the flow of
rename.» Issue открыт по сей день.
— Общая позиция: «The rename refactoring for example added an additional `prepareRename` request that
ask the server to do some validation and let the client know about it. The client can then decide what
to do (e.g. can even show some UI) and then execute `rename` or do nothing. If we need such a multi step
flow in other scenarios then we should look into those scenarios and come up with a data driven protocol
for them.» — `LSPI1641`, dbaeumer, 2023-01-16.
— «I am actually against this. Knowing what I know today I would not add `showMessage` and
`showMessageRequest`. IMO a client needs to provide a reasonable set of input to a request and a server
shouldn't have the need to ask for additional input. This ensure that servers can work without UI.» —
`LSPI1641`, dbaeumer, 2023-01-15.

**Тезис 1.6. Запрос «уточнить семантику ошибок rename» был закрыт без изменений спецификации.**
— `LSPI566` «Error handling in Rename» (2018-09-08): вопросы «what error code should be used if
`newName` is invalid?», «what is the valid use of `null`?». dbaeumer 2018-09-12: «I am all in for
clarifying the spec here. As always PRs are welcome.» Закрыт 2021-10-28: «The feature has not gain any
traction in the community. I therefore close the issue.»
— `LSPI800` (clangd, 2019-07-23): просьба закрепить в спеке, что клиент показывает message из error
для `prepareRename`. dbaeumer: «The message of the response error will then make it into the UI.» —
формально в спеке только фраза «Clients should show the information in their user interface».

---

## 2. Какие рефакторинги — операции протокола, а какие — только `codeAction`

**Тезис 2.1. Единственный рефакторинг, являющийся первоклассной операцией протокола, — rename (символа) плюс файловые операции `workspace/willRenameFiles`. Change Signature, Move, Safe Delete, Inline, Extract как операции протокола не существуют.**
— Полный список методов раздела Language Features в 3.17 (по анкорам страницы): declaration, definition,
typeDefinition, implementation, references, prepareCallHierarchy/incomingCalls/outgoingCalls,
prepareTypeHierarchy/supertypes/subtypes, documentHighlight, documentLink, hover, codeLens,
foldingRange, selectionRange, documentSymbol, semanticTokens, inlayHint, inlineValue, moniker,
completion, publishDiagnostics/pull diagnostics, signatureHelp, codeAction, documentColor, formatting,
rangeFormatting, onTypeFormatting, **rename, prepareRename**, linkedEditingRange.
— `SPEC317#languageFeatures`
— gopls формулирует то же самое (иллюстрация): «Code transformations are not a single category in the
LSP: A few, such as Formatting and Rename, are primary operations in the protocol. … Most transformations
are defined as code actions.» — https://go.dev/gopls/features/transformation
— Ответ мейнтейнера на самый первый запрос операции refactoring: «no there isn't. We started a
discussion about how a refactoring protocol could look like but we haven't found a good way to spec that
yet since this is very dependent on the programming language you are using.» — `LSPI61`, dbaeumer,
2016-09-19. Issue открыт с 2016 года, последний вопрос «Is this still being looked at?» 2023-12-22
без ответа.

**Тезис 2.2. `CodeActionKind` — открытый набор строк-категорий; предопределённые `refactor.extract`, `refactor.inline`, `refactor.rewrite` (3.17) и `refactor.move` (3.18) — это только ярлыки для группировки в меню, без параметров и семантики.**
— «Kinds are a hierarchical list of identifiers separated by `.`, e.g. `"refactor.extract.function"`.
The set of kinds is open and client needs to announce the kinds it supports to the server during
initialization.»
— Примеры в комментариях: `refactor.extract`: «Extract method, Extract function, Extract variable,
Extract interface from class»; `refactor.inline`: «Inline function, Inline variable, Inline constant»;
`refactor.rewrite`: «Convert JavaScript function to class, Add or remove parameter, Encapsulate field,
Make method static, Move method to base class».
— «the ability to group code actions using a kind. Clients are allowed to ignore that information.
However it allows them to better group code action for example into corresponding menus».
— `SPEC317#codeActionKind`, `SPEC317#textDocument_codeAction`
— 3.18: «Base kind for refactoring move actions: 'refactor.move' … Move a function to a new file, Move a
property between classes, Move method to base class» — `SPEC318#codeActionKind`.
— Заметьте: «Add or remove parameter» (то есть Change Signature) приведён как пример `refactor.rewrite`,
но никакого способа передать новую сигнатуру в протоколе нет (см. 2.3).

**Тезис 2.3. Code action не принимает параметров от клиента. Вход — только `textDocument`, `range`, `context` (диагностики, фильтр kinds, triggerKind). Выход — `edit` и/или `command` с аргументами, которые сервер сам заранее заполнил.**
— `CodeActionParams`: «textDocument … range … context: CodeActionContext» — `SPEC317#codeActionParams`.
— «A CodeAction must set either `edit` and/or a `command`. If both are supplied, the `edit` is applied
first, then the `command` is executed.» — `SPEC317#codeAction`.
— `Command.arguments?: LSPAny[]` — «Arguments that the command handler should be invoked with.»; «The
protocol currently doesn't specify a set of well-known commands.» — `SPEC317#command`.
— `workspace/executeCommand`: «The arguments are typically specified when a command is returned from
the server to the client.» — `SPEC317#workspace_executeCommand`. То есть клиент может подменить
`arguments`, но что в них лежит — приватный контракт конкретного сервера, не протокола.
— Единственная структурная «причина» у code action — `disabled.reason: string` («Human readable
description of why the code action is currently disabled»), опять строка для UI — `SPEC317#codeAction`.

**Тезис 2.4. Запроса «сервер просит у клиента ввод» в протоколе нет; `window/showMessageRequest` позволяет только выбрать один из предложенных сервером пунктов, `window/showDocument` — только открыть URI.**
— `showMessageRequest`: «the request allows to pass actions and to wait for an answer from the client»;
`MessageActionItem { title: string }`; «result: the selected MessageActionItem | null if none got
selected.» — `SPEC317#window_showMessageRequest`. Ввести строку (имя, путь) через него нельзя.
— `showDocument`: «ask the client to display a particular resource referenced by a URI in the user
interface.» — `SPEC317#window_showDocument`.
— Запрос ввода строки (`LSPI1641`, 2023-01-13, по мотивам Metals `metals/inputBox`) отклонён мейнтейнером
(цитаты в 1.5). Запрос модального диалога (`LSPI1337`, 2021-08-26, Pylance): «I personally think it
is a bad idea to let servers open modal dialogs.» (dbaeumer 2021-08-27); «LSP is about standardizing
programming language messages and not UI. I would like to leave the UI part to the client.» (dbaeumer
2022-05-23). Открыт.
— Для агента: `showMessageRequest` формально можно обработать программно (агент «нажимает» кнопку), но
спека не гарантирует, что сервер вообще его пришлёт, и не даёт ему структуры.

**Тезис 2.5. Рефакторинги с параметрами (Move to file, Extract с именем, Change Signature) — открытая проблема с 2020 года; в сентябре 2026 команды Go и Dart опубликовали предложение `command/resolve` с формами, но в спеку (3.18, черновик 3.19) оно не вошло.**
— `LSPI1164` «Support refactors that require user input/options» (DanTup, 2020-12-09): «One common
request of Dart is a refactor for "move to file" … there's no way to ask the user to pick a file».
— Pylance, 2022-11-04 (heejaechang): «We want to implement `rename file code action`, which needs user
input for the file name. We want to implement `move symbol`, which requires `file name` … 90% of them
require very simple user input, yet we have to have custom code on the client side».
— Предложение Go/Dart, 2026-09-10 (h9jiang): новый запрос `command/resolve` с `InteractiveParams
{ formFields, formAnswers }`, stateless по модели HTML-форм; «**No support for edit-based CodeActions**:
Interactivity is not supported for CodeActions that resolve directly to `WorkspaceEdit` payloads.»
Прототипы — под `experimental` в gopls. Ссылка на дизайн:
https://go.googlesource.com/proposal/+/refs/heads/master/design/76331-lsp-interactive-refactoring.md
— Issue открыт; в `SPEC318` и черновике 3.19 слов `interactive`/`command/resolve`/`formField` нет
(проверено grep'ом по тексту спецификаций).

**Тезис 2.6. Что 3.18 добавило по теме: `SnippetTextEdit` (для «extract → сразу переименовать»), `refactor.move`, `CodeActionTag.LLMGenerated`, `WorkspaceEditMetadata.isRefactoring`.**
— «Since 3.18.0, there is also the concept of a snippet text edit, which supports inserting a snippet
instead of plain text.» Но: «interactive snippets are only applied to the file opened in the active
editor» и «In case the snippet text edit corresponds to a file that is not currently open in the active
editor, the client should downgrade the snippet to a non-interactive normal text edit» —
`SPEC318#snippetTextEdit`. Для агента без «active editor» это просто текст с плейсхолдерами.
— История: `LSPI764` «How can a server implement "extract method" (how can it trigger rename?)»
(2019-05-29) → dbaeumer 2024-07-08: «Snippers made it into code actions and workspace edits for 3.18».
— `CodeActionTag`: «Marks the code action as LLM-generated.» (`LLMGenerated = 1`) — `SPEC318#codeActionTag`.
— `WorkspaceEditMetadata { isRefactoring?: boolean }` — «Signal to the editor that this edit is a
refactoring.» — `SPEC318#workspaceEditMetadata`. Только флаг для UI (в VS Code управляет «рефакторингом»
для авто-сохранения/preview), никакой семантики.

---

## 3. Применение `WorkspaceEdit`: кто применяет, атомарность, версии, файлы вне клиента

**Тезис 3.1. Правки применяет клиент. Сервер только описывает изменение (в rename/codeAction) или просит применить (`workspace/applyEdit`).**
— «The rename request is sent from the client to the server to ask the server to compute a workspace
change so that the client can perform a workspace-wide rename of a symbol.» — `SPEC317#textDocument_rename`.
— «The workspace/applyEdit request is sent from the server to the client to modify resource on the client
side.» — `SPEC317#workspace_applyEdit`.
— «In most cases the server creates a WorkspaceEdit structure and applies the changes to the workspace
using the request workspace/applyEdit» — `SPEC317#workspace_executeCommand`.

**Тезис 3.2. Атомарность — не гарантия протокола, а самодекларация клиента через `failureHandling`. Есть четыре режима, включая «abort» (частично применённые правки остаются) и «undo» («no guarantee»).**
— «`abort`: Applying the workspace change is simply aborted if one of the changes provided fails. All
operations executed before the failing operation stay executed.»
— «`transactional`: All operations are executed transactional. That means they either all succeed or no
changes at all are applied to the workspace.»
— «`textOnlyTransactional`: If the workspace edit contains only textual file changes they are executed
transactional. If resource changes (create, rename or delete file) are part of the change the failure
handling strategy is abort.»
— «`undo`: The client tries to undo the operations already executed. But there is no guarantee that this
is succeeding.»
— `SPEC317#failureHandlingKind`, `SPEC317#workspaceEditClientCapabilities`
— Следствие для агента-клиента: агент сам обязан реализовать транзакционность и сам объявляет
`failureHandling`; сервер этому только верит. Вывод: «атомарный рефакторинг» — свойство клиента, не
протокола.
— Иллюстрация: `vscode-languageclient` объявляет `FailureHandlingKind.TextOnlyTransactional`
(https://github.com/microsoft/vscode-languageserver-node/blob/main/client/src/common/client.ts,
строка `workspaceEdit.failureHandling = FailureHandlingKind.TextOnlyTransactional;`), то есть даже
эталонный клиент не даёт транзакционности при create/rename/delete файлов.

**Тезис 3.3. Ответ `applyEdit` — булево `applied`, опциональная строка `failureReason` и опциональный индекс `failedChange`. Индекс — только если клиент заявил `failureHandling`.**
— «`applied: boolean` Indicates whether the edit was applied or not.»
— «`failureReason?: string` An optional textual description for why the edit was not applied. This may
be used by the server for diagnostic logging or to provide a suitable error for a request that triggered
the edit.»
— «`failedChange?: uinteger` Depending on the client's failure handling strategy `failedChange` might
contain the index of the change that failed. This property is only available if the client signals a
`failureHandling` strategy in its client capabilities.»
— `SPEC317#applyWorkspaceEditResult`
— Обратной связи «рефакторинг не удался/неполон» от сервера клиенту в этом потоке нет: `LSPI1988`
«Signal the client a refactoring failed» (2024-07-22): «the `ApplyWorkspaceEditParams` type does not
contain anything to report a problem that occured during the edits computation … the `showMessage`
notification is not appropriate because the client has no info to correlate an incoming `showMessage`
with a previously launched refactoring.» Открыт, без ответа мейнтейнеров.

**Тезис 3.4. Порядок операций: ресурсные операции — строго по порядку; текстовые правки внутри одного документа — на одной версии, без перекрытий.**
— «If resource operations are present clients need to execute the operations in the order in which they
are provided. … An invalid sequence (e.g. (1) delete file a.txt and (2) insert text into file a.txt) will
cause failure of the operation. How the client recovers from the failure is described by the client
capability: workspace.workspaceEdit.failureHandling» — `SPEC317#workspaceEdit`.
— «All text edits ranges refer to positions in the document they are computed on. They therefore move a
document from state S1 to S2 without describing any intermediate state. Text edits ranges must never
overlap» — `SPEC317#textEditArray`.
— «A TextDocumentEdit describes all changes on a version Si and after they are applied move the document
to version Si+1.» — `SPEC317#textDocumentEdit`.

**Тезис 3.5. Версии документов: сервер может указать версию (`OptionalVersionedTextDocumentIdentifier`), для неоткрытых файлов — `null` («мастер — диск»). Что делать при несовпадении версии, спека не говорит.**
— «The version number of this document. If an optional versioned text document identifier is sent from
the server to the client and the file is not open in the editor (the server has not received an open
notification before) the server can send `null` to indicate that the version is known and the content on
disk is the master (as specified with document content ownership).» —
`SPEC317#optionalVersionedTextDocumentIdentifier`.
— «The text document is referred to as a OptionalVersionedTextDocumentIdentifier to allow clients to
check the text document version before an edit is applied.» — `SPEC317#textDocumentEdit`. «to allow
clients to check» — возможность, не обязанность; реакция на расхождение не специфицирована.
— Используется только если клиент заявил `workspace.workspaceEdit.documentChanges`; иначе — плоский
`changes: { [uri]: TextEdit[] }` без версий вовсе — `SPEC317#workspaceEdit`.
— Иллюстрация (клиент): `vscode-languageclient` при несовпадении версии молча отвечает
`{ applied: false }` без `failureReason` (`validateWorkspaceEdit` в `client/src/common/client.ts`):
«If not, we can't apply the workspace edit since it might be based on an old version of the document.»

**Тезис 3.6. Файлы вне клиента: «открыт» в LSP = «клиент владеет содержимым», а не «виден в редакторе». Для неоткрытых файлов мастер — диск, и сервер вправе править их сам.**
— «The document's content is now managed by the client and the server must not try to read the
document's content using the document's Uri. Open in this sense means it is managed by the client. It
doesn't necessarily mean that its content is presented in an editor.» — `SPEC317#textDocument_didOpen`.
— «The document's master now exists where the document's Uri points to … Note that a server's ability to
fulfill requests is independent of whether a text document is open or closed.» — didClose, там же.
— dbaeumer о том, как сервер узнаёт об изменениях после applyEdit: «for all resources manipulated via a
workspace edit we open them in an internal buffer … the client changes the file … the client closes the
file … So the bottom line is that the server content should be updated without that the server replays the
edits. If the server directly manipulates the files on disk (which he is allowed to do for files which
are not open) then the server is responsible to update its internal buffers as well.» — `LSPI61`,
2017-01-11. Это описание поведения VS Code, в спеке этого протокола-обязательства нет.
— Для агента: если агент правит файлы на диске сам (не через didOpen/didChange), сервер узнаёт об этом
только через `workspace/didChangeWatchedFiles`, которые должен слать клиент: «The watched files
notification is sent from the client to the server when the client detects changes to files and folders
watched by the language client» — `SPEC317#workspace_didChangeWatchedFiles`.

**Тезис 3.7. Устаревшие результаты: спека даёт паттерн (клиент отменяет/пересылает; сервер может ответить `ContentModified`), но не гарантию.**
— «We recommend the following implementation pattern to avoid that clients apply outdated response
results: if a client sends a request to the server and the client state changes in a way that it
invalidates the response it should … cancel the server request and ignore the result … if a server
detects an internal state change (for example, a project context changed) that invalidates the result of
a request in execution the server can error these requests with ContentModified.» —
`SPEC317#implementationConsiderations`.
— `ContentModified = -32801`: «The server detected that the content of a document got modified outside
normal conditions. A server should NOT send this error code if it detects a content change in its
unprocessed messages. The result even computed on an older state might still be useful for the client.»
— `SPEC317#errorCodes`.
— Параллельные rename — не определено: «I think there isn't a clear answer to this from a server
perspective since it depends on the client capabilities … a client shouldn't do this unless it can
handle both results appropriately.» — `LSPI1663`, dbaeumer, 2023-02-14.

---

## 4. Готовность и индексация

**Тезис 4.1. Сигнала «workspace проиндексирован, результаты полны» в протоколе нет. Есть только прогресс — индикатор для UI, без семантики.**
— `WorkDoneProgressBegin.title`: «Mandatory title of the progress operation. Used to briefly inform
about the kind of operation being performed. Examples: "Indexing" or "Linking dependencies".»
— `message`: «Examples: "3/25 files", "project/src/module2", "node_modules/some_dep".»
— `percentage`: «If not provided infinite progress is assumed».
— `WorkDoneProgressEnd.message`: «Optional, a final message indicating to for example indicate the
outcome of the operation.»
— `SPEC317#workDoneProgressBegin`, `#workDoneProgressEnd`
— Ни одно из полей не машиночитаемо: "Indexing" — просто строка-пример, а не зарезервированное значение.

**Тезис 4.2. Серверный прогресс (`window/workDoneProgress/create`) возможен только если клиент заявил `window.workDoneProgress`; иначе сервер не вправе его слать.**
— «To keep the protocol backwards compatible servers are only allowed to use
window/workDoneProgress/create request if the client signals corresponding support using the client
capability window.workDoneProgress» — `SPEC317#serverInitiatedProgress`.
— «This is useful if the server needs to report progress outside of a request (for example the server
needs to re-index a database).» — там же.

**Тезис 4.3. После ответа на `initialize` сервер считается готовым принимать запросы; «готов» ≠ «проиндексирован». Мейнтейнер отказался вводить readiness-сигнал.**
— «Until the server has responded to the initialize request with an InitializeResult, the client must
not send any additional requests or notifications to the server.» — `SPEC317#initialize`. Обратного
условия («после ответа результаты полны») в спеке нет.
— `LSPI511` «Discussion: LSP-server readiness indicator» (2018-06-29), открыт. dbaeumer 2020-07-02:
«If a server receives a request and it is not ready to serve the request IMO the best to do is that
clients setup progress tokens which then can be used by the server to reports this. In an async world I
don't see how a ready notification / request from the server can make this fail save.»
— puremourning (YouCompleteMe) 2020-07-02: «most servers simply reply immediately to the initialise
response then use some out-of-band mechanism to report readiness … the server is 'initialised' and will
process requests after the initialise response, but it's not really "fully ready" until all the indexes
are up to date.»
— robertoaloi (Erlang LS) 2020-07-02: «What would be much better is for the protocol to allow us to
notify that we are still not "fully ready".»
— matklad (rust-analyzer) 2021-04-06: «we employ a simple hack of answering with `ContentModified`, which
causes the client to ignore the error … the server refuses to respond because it knows that the response
will be immediately obsoleted by indexing results.»

**Тезис 4.4. Индексация как таковая сознательно оставлена вне протокола.**
— `LSPI54` «Clarification for the Indexing workflow» (2016-08-23): просьба специфицировать «How the
server notifies the client that the index is complete». dbaeumer 2016-09-17: «the protocol should not
define how indexing works. This will be very hard.» Закрыт 2019-10-30: «Version 3.15 will have support
for progress reporting. I will close the issue as out of scope.»

**Тезис 4.5. К наблюдению «pyright: 2 ссылки сразу после initialize, 79 через 40 с» протокол относится так: это поведение, которое спека не запрещает и не описывает; ответ `Location[]` не несёт признака полноты.**
— В разделе `textDocument/references` нет ни слова о полноте (см. раздел 6).
— Единственная оговорка о неполноте есть только для отменённых запросов с partial results: «the code
equals to RequestCancelled: the client is free to use the provided results but should make clear that the
request got canceled and may be incomplete.» — `SPEC317#partialResults`.
— Иллюстрация (pyright, не протокол): `microsoft/pyright#10086` (2025-03-14) «Pyright LSP Returns
Incomplete References Without Explicit didOpen for All Files»: «When I send didOpen only for the file
where I make the references request, Pyright returns just 3 references … in debug mode—waiting a few
seconds … it returns 102 references». Причина по ответу rchiodo (Pylance) 2025-03-17: «I believe
'rootUri' is deprecated. We don't use it anymore. You have to send workspace folders.» Спека: `rootUri`
— «@deprecated in favour of `workspaceFolders`» (`SPEC317#initializeParams`). Отдельно erictraut
2025-03-15: «Pyright is focused on type checking. While it does have some basic language server support,
the features are quite limited. Pylance … provides more advanced language server features including
indexing.»

---

## 5. Мультиязычность: один сервер — один язык

**Тезис 5.1. Модель протокола: инструмент запускает по серверу на язык; композиции между серверами в протоколе нет.**
— «When a user is working with different languages, a development tool usually starts a language server
for each programming language.» — https://microsoft.github.io/language-server-protocol/overviews/lsp/overview/
— «The data types are not at the level of a programming language domain model which would usually provide
abstract syntax trees and compiler symbols … It is much simpler to standardize a text document URI or a
cursor position compared with standardizing an abstract syntax tree and compiler symbols across different
programming languages.» — там же.
— «There is no generic composition mechanism in the LSP.» — `LSPI636`, dbaeumer, 2018-12-18.
— «I think LSP should not promote a model how to do this. I think both using forwards or embedded
services is a valid solution.» — `LSPI636`/`LSPI1252`, dbaeumer, 2021-10-28.
— Про лишние серверы и вотчеры: «a client usually starts more than one server. If every server runs its
own file system watching it can become a CPU or memory problem.» — `SPEC317#workspace_didChangeWatchedFiles`.

**Тезис 5.2. Кросс-языковых ссылок и кросс-языкового rename в протоколе нет; слияние результатов нескольких серверов — дело клиента, спека его не описывает.**
— В спеке слова «merge»/«multiple servers» применительно к результатам отсутствуют (grep по тексту 3.17
и 3.18).
— Иллюстрация — документированное поведение VS Code (API `vscode.d.ts`):
  - references: «Multiple providers can be registered for a language. In that case providers are asked
    in parallel and the results are merged. A failing provider (rejected promise or exception) will not
    cause a failure of the whole operation.» — `languages.registerReferenceProvider`.
  - rename: «Multiple providers can be registered for a language. In that case providers are sorted by
    their {@link languages.match score} and asked in sequence. The first provider producing a result
    defines the result of the whole operation.» — `languages.registerRenameProvider`.
    https://github.com/microsoft/vscode/blob/main/src/vscode-dts/vscode.d.ts
  - Реализация: `src/vs/editor/contrib/rename/browser/rename.ts` (`RenameSkeleton._provideRenameEdits`
    — перебор провайдеров до первого не-null результата, отказы копятся в `rejectReason`).
— То есть даже в эталонном клиенте rename никогда не объединяет правки двух серверов (например,
Kotlin + Java в одном проекте): побеждает первый ответивший.

---

## 6. Полнота `textDocument/references`

**Тезис 6.1. Спека обещает «project-wide references», параметр `includeDeclaration`, результат `Location[] | null` — и ничего о полноте, покрытии неоткрытых файлов, зависимостей или динамических вызовов.**
— «The references request is sent from the client to the server to resolve project-wide references for
the symbol denoted by the given text document position.»
— `ReferenceContext.includeDeclaration: boolean` — «Include the declaration of the current symbol.»
— «result: `Location[] | null`; partial result: `Location[]`; error: code and message set in case an
exception happens during the reference request.»
— `SPEC317#textDocument_references`
— В `Location` нет типа ссылки (чтение/запись/комментарий/неподтверждённая): `LSPI763` «Add "type" to
textDocument/references results» (2019-05-28, C/C++ extension: «Confirmed reference … Not a reference
(i.e. has the same name, but fails semantic checking)») — открыт.

**Тезис 6.2. Partial results — стриминг, а не признак полноты. Если сервер стримит, финальный ответ пуст; отменённый стрим «may be incomplete».**
— «If a server reports partial result via a corresponding $/progress, the whole result must be reported
using n $/progress notifications. Each of the n $/progress notification appends items to the result. The
final response has to be empty in terms of result values.»
— «If the response errors the provided partial results should be treated as follows: the code equals to
RequestCancelled: the client is free to use the provided results but should make clear that the request
got canceled and may be incomplete. in all other cases the provided partial results shouldn't be used.»
— `SPEC317#partialResults`
— Клиент включает это только присылая `partialResultToken` («Whether a client accepts partial result
notifications for a request is signaled by adding a partialResultToken to the request parameter»).

---

## 7. Прочие возможности, полезные агенту (кратко, только как capabilities)

- `workspace/symbol`: «list project-wide symbols matching the query string»; «Clients may send an empty
  string here to request all symbols.» Лимитов/пагинации нет; с 3.17 — `workspaceSymbol/resolve` для
  ленивого range. — `SPEC317#workspace_symbol`.
- `textDocument/implementation` (3.6): «resolve the implementation location of a symbol at a given text
  document position.» — `SPEC317#textDocument_implementation`.
- Call hierarchy (3.16): двухшаговый — `textDocument/prepareCallHierarchy` → `callHierarchy/incomingCalls`
  / `outgoingCalls`. — `SPEC317#textDocument_prepareCallHierarchy`.
- Type hierarchy (3.17): `textDocument/prepareTypeHierarchy` → `typeHierarchy/supertypes` / `subtypes`;
  «Will return null if the server couldn't infer a valid type from the position.» —
  `SPEC317#textDocument_prepareTypeHierarchy`.
- `textDocument/linkedEditingRange` (3.16): синхронная правка одинаковых диапазонов — текстовая, не
  семантическая. — `SPEC317#textDocument_linkedEditingRange`.
- `workspace/willRenameFiles` (3.16): сервер может вернуть `WorkspaceEdit` на переименование файла —
  единственный «move»-подобный первоклассный запрос, и он про файлы, не про символы. —
  `SPEC317#workspace_willRenameFiles`.
- Все запросы позиционные: «It is up to the client to decide how a selection is converted into a
  position» — `SPEC317#textDocumentPositionParams`. Адресовать символ по имени/FQN нельзя, только по
  (uri, line, character).
- `general.staleRequestSupport` (3.17): клиент перечисляет, какие запросы он повторит при
  `ContentModified` («retryOnContentModified: string[]») — `SPEC317#initializeParams`.

---

## Что можно говорить со сцены на уровне протокола

1. **В LSP ровно один семантический рефакторинг — rename символа. Всё остальное — `codeAction` с
   ярлыком `refactor.*`, без параметров.** Источник: список методов Language Features (`SPEC317`),
   `CodeActionParams`/`CodeAction` (`SPEC317#textDocument_codeAction`), `LSPI61` (dbaeumer 2016-09-19:
   «no there isn't»).
2. **Ответ rename — либо `WorkspaceEdit`, либо ошибка с текстом. Структурного «есть конфликт, но можно
   продолжить» нет.** Источник: `SPEC317#textDocument_rename` («error: code and message set in case when
   rename could not be performed for any reason»).
3. **`prepareRename` не видит нового имени — он вызывается до его ввода и физически не может проверить
   конфликт.** Источник: `PrepareRenameParams` (`SPEC317#textDocument_prepareRename`), dbaeumer
   `LSPI566` 2018-09-12.
4. **Единственный протокольный способ предупредить — `ChangeAnnotation.needsConfirmation` на правках
   (3.16), и сервер вправе слать его только клиенту, заявившему `changeAnnotationSupport` /
   `honorsChangeAnnotations`. Это строка для человека, не код для машины.** Источник:
   `SPEC317#changeAnnotation`, `#renameClientCapabilities`, `#textDocumentEdit`.
5. **Сервер не может попросить ввод. Мейнтейнеры отклонили input-request (2023) и модальные диалоги
   (2021); задача «рефакторинг с параметрами» открыта с 2020, предложение Go/Dart (`command/resolve`,
   сентябрь 2026) пока вне спецификации.** Источник: `LSPI1641` (dbaeumer 2023-01-15/16), `LSPI1337`
   (dbaeumer 2021-08-27, 2022-05-23), `LSPI1164` (h9jiang 2026-09-10), `SPEC318`/3.19 без таких методов.
6. **Атомарность применения правок — не гарантия протокола, а самодекларация клиента:
   `failureHandling` ∈ {abort, transactional, textOnlyTransactional, undo}, причём «undo … no
   guarantee».** Источник: `SPEC317#failureHandlingKind`. Обратной связи «рефакторинг провалился/неполон»
   от сервера нет — `LSPI1988` (2024, открыт).
7. **Сигнала «индекс готов, результаты полны» в протоколе нет; `$/progress` с title "Indexing" — пример
   строки для UI. Спека `references` не говорит о полноте ни слова.** Источник:
   `SPEC317#workDoneProgressBegin`, `#textDocument_references`, `LSPI511` (dbaeumer 2020-07-02),
   `LSPI54` (закрыт как out of scope, 2019-10-30).
8. **Один сервер — один язык; композиции и кросс-языкового rename в протоколе нет («There is no generic
   composition mechanism in the LSP»).** Источник: overview LSP, `LSPI636` (dbaeumer 2018-12-18,
   2021-10-28).

## Что НЕЛЬЗЯ утверждать (не подтверждено спекой или зависит от сервера/клиента)

- «LSP запрещает серверу предупреждать о конфликтах» — неверно: механизм `ChangeAnnotation` есть, просто
  он необязателен, текстовый и рассчитан на UI. Правильно: «нет структурного, обязательного канала».
- «Rename в LSP всегда молча ломает код при коллизии имён» — это поведение конкретного сервера
  (см. иллюстрации: clangd и gopls отказывают, rust-analyzer аннотирует, jdtls глотает warnings).
  Протокол ни того, ни другого не требует.
- «`prepareRename` возвращает причину отказа» — только `message` строкой в error; поле причины не
  специфицировано. И VS Code показывает этот текст — это поведение VS Code (`LSPI800`), спека говорит
  лишь «Clients should show the information in their user interface».
- «references после initialize неполны, потому что так устроен LSP» — спека не описывает индексацию
  вовсе; неполнота — свойство сервера (pyright: `#10086`, плюс `rootUri` deprecated). Говорить можно
  только «протокол не даёт сигнала полноты».
- «LSP гарантирует транзакционность WorkspaceEdit» — нет; зависит от заявленного клиентом
  `failureHandling`, и даже VS Code заявляет только `textOnlyTransactional`.
- «Сервер проверяет версию документа перед правкой» — версии в `TextDocumentEdit` опциональны и есть
  только при `documentChanges`; проверка — на стороне клиента и не обязательна («to allow clients to
  check»).
- «В LSP есть Change Signature / Move / Inline / Safe Delete» — есть только строки-категории
  `refactor.rewrite` / `refactor.move` (3.18) / `refactor.inline` для меню; операций с параметрами нет.
  «Add or remove parameter» в спеке — пример названия, не операция.
- «3.18 решил проблему интерактивных рефакторингов» — 3.18 добавил только `SnippetTextEdit`
  (плейсхолдеры в тексте, понижаемые до обычного текста для неактивных файлов), `refactor.move`,
  `CodeActionTag.LLMGenerated` и `WorkspaceEditMetadata.isRefactoring`.
- «Можно спросить у сервера, готов ли он» — нет такого запроса; rust-analyzer/clangd/jdtls используют
  собственные расширения (`LSPI511`).
- Любые цифры («79 ссылок через 40 с», «2 ссылки сразу») — только как наблюдение за конкретным сервером
  в конкретной конфигурации.

## Иллюстрации (server-specific, помечать как таковые)

1. **clangd** — при конфликте имени rename отвечает ошибкой, продолжить нельзя: `InvalidName::Conflict`
   → `"conflict with the symbol in {0}"` → `error("invalid name: {0}", …)`. Функции/методы сознательно
   не проверяются: «Function conflicts are subtle (overloading), so ignore them.»
   https://github.com/llvm/llvm-project/blob/main/clang-tools-extra/clangd/refactor/Rename.cpp
   (`checkName`, `makeError(InvalidName)`).
2. **gopls** — документированные проверки при rename, все завершаются отказом: «Gopls' renaming algorithm
   takes great care to detect situations in which renaming might introduce a compilation error. For
   example, changing a name may cause a symbol to become "shadowed" … Gopls will report an error, stating
   the pair of symbols and the shadowed reference … it aborts the renaming with an error.» Там же: «(This
   simple dialog support is unique among LSP refactoring operations; see
   microsoft/language-server-protocol#1164.)» https://go.dev/gopls/features/transformation#rename
3. **rust-analyzer** — предупреждает, но даёт продолжить, через протокольный `ChangeAnnotation`:
   `label: "This rename will change the program's meaning"`, `needs_confirmation: true`, `description:
   "Some variable(s) will shadow the renamed variable or be shadowed by it if the rename is performed"`
   (`crates/ide-db/src/rename.rs`). Автор PR о механизме: «We do it by introducing a dummy edit with an
   annotation. I'm not a fond of this approach, but I don't think LSP has a better way.» —
   https://github.com/rust-lang/rust-analyzer/pull/19079 (2025-02-01, merged 2025-03-10). Настройка
   `rust-analyzer.rename.showConflicts` (PR #20193, 2025-07-07). Клиент без `changeAnnotationSupport`
   этого предупреждения не увидит.
4. **jdtls (Eclipse JDT LS)** — четырёхуровневый `RefactoringStatus` Eclipse (INFO/WARNING/ERROR/FATAL)
   схлопывается до «FATAL → ResponseError, иначе молча вернуть правки»:
   `if (check.getStatus().getSeverity() >= RefactoringStatus.FATAL) throw new ResponseErrorException(new
   ResponseError(ResponseErrorCode.InvalidRequest, check.getStatus().getMessageMatchingSeverity(
   RefactoringStatus.ERROR), null));` — предупреждения уровня WARNING/ERROR теряются.
   https://github.com/eclipse-jdtls/eclipse.jdt.ls/blob/master/org.eclipse.jdt.ls.core/src/org/eclipse/jdt/ls/core/internal/handlers/RenameHandler.java
5. **JetBrains Kotlin LSP** — наблюдение 2026-09-29 (не документ): rename в существующее имя отвергается
   ошибкой «Function 'lint' is already declared». Сервер «based on IntelliJ IDEA and the IntelliJ IDEA
   Kotlin Plugin» (README https://github.com/Kotlin/kotlin-lsp), то есть это IntelliJ-овская проверка
   конфликтов, доставленная через единственный доступный канал — error rename.
6. **pyright** — неполные references до «прогрева»/без workspaceFolders: `microsoft/pyright#10086`
   (2025-03-14), см. тезис 4.5. Плюс историческое `#375` (2019): «"Rename Symbol" functionality uses the
   same code paths as "Find All References"» (erictraut) — то есть неполнота references напрямую даёт
   неполный rename.
7. **VS Code как клиент** — rename берёт результат первого провайдера, references объединяет всех
   (`vscode.d.ts`); `vscode-languageclient` заявляет `textOnlyTransactional` и при расхождении версии
   отвечает `{ applied: false }` без причины (`client/src/common/client.ts`).
8. **Dart** (не протокол, обходной путь) — интерактивные рефакторинги через приватный контракт в
   `CodeAction.data.parameters` и обёртку команды на стороне VS Code-расширения: «There's no spec for
   this yet» — `LSPI1164`, DanTup, 2022-11-28.
