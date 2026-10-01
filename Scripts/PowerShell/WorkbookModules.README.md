# Workbook module installation

`Update-WorkbookModules.ps1` is the common entry point for installing VBA sources into PERSONAL, application workbooks and WorkbookUpdater. `WorkbookModules.ps1` implements schema resolution, source preparation and component replacement. PERSONAL command launchers, the updater build wrapper and the legacy closed-workbook reload wrapper all use this implementation.

## Source schema

Pass `-VbaFolderPath <folder>` containing `modules.json`. Without the parameter, the folder is `vba` beside `Update-WorkbookModules.ps1`, independently of the current working directory. Schema paths are relative to that folder.

The schema maps profile names to nonempty arrays of source paths or wildcard patterns. Use `-Profile <name>`, or omit it for an existing workbook with `tbConfig` containing `ThisWorkbook::id`. There is no implicit Default profile. Literal bracket folder names use the existing `[[]` syntax, for example `common/[[]0] interfaces/*.cls.vba`. Unmatched patterns, invalid paths, empty files and duplicate component names are rejected before replacement.

UTF-8 `.vba`, `.cls.vba` and `.cls.utf8.vba` sources are prepared for `CodeModule.AddFromString`: class export headers and attributes are removed, and Unicode string literals use `VBA.ChrW$`. Native `.bas`, `.cls`, `.frm` and associated `.frx` files are imported as byte snapshots in their original Excel export encoding; native attributes are preserved. UserForms require native exports.

`ThisWorkbook.vba` targets the actual workbook CodeName. `ws_<sheet name>.vba` targets an existing worksheet by its display name and writes to its CodeName. The profile defines the complete project: standard modules, classes and forms are removed, and all existing document code is cleared, including code omitted from the profile. Clear mode performs the same removal without importing replacement sources.

## Commands

Preview a source plan without Excel:

```powershell
.\Scripts\PowerShell\Update-WorkbookModules.ps1 -VbaFolderPath .\PROJECTS_EXCEL\vba -Profile PersonalEventBuilder -PlanOnly
```

Update loaded PERSONAL and initialize its new runtime:

```powershell
.\Scripts\PowerShell\Update-WorkbookModules.ps1 -WorkbookName PERSONAL.XLSB -VbaFolderPath .\MacrosExcel\PERSONAL -Profile PERSONAL -InitializeMacro ex_Core.fn_ReloadPersonalRuntime
```

Update an existing, unloaded WorkbookUpdater without rebuilding the file:

```powershell
.\Scripts\PowerShell\Update-WorkbookModules.ps1 -WorkbookPath .\PROJECTS_EXCEL\WorkbookUpdater\WorkbookUpdater.xlam -VbaFolderPath .\PROJECTS_EXCEL\WorkbookUpdater -Profile WorkbookUpdater
```

Build a new WorkbookUpdater using the same schema and importer:

```powershell
.\PROJECTS_EXCEL\WorkbookUpdater\Build-Updater.ps1 -OutputPath C:\Temp\WorkbookUpdater.xlam
```

Create other workbooks with `-Create -WorkbookPath <new.xlsm>`. `-AsAddin` creates `.xlam` files. Creation refuses to overwrite an existing path. `-WorkbookName` targets a named workbook in the active Excel instance; `-WorkbookPath` uses a separate hidden instance with events and automatic macros disabled. It does not reopen the result in the user instance.

## Lifecycle and recovery

This installer requires the target VBA project to be in Design mode. It does not implement cooperative stopping of arbitrary running application code. `-CanUpdateMacro` optionally calls a Boolean permission check; `-PrepareMacro` invokes an explicit cleanup procedure after backup and before replacement. `-InitializeMacro` runs after replacement and before saving. Callback names are qualified to the exact target workbook.

For WorkbookUpdater, first complete or cancel the pending operation, ensure its target books are Running, and unload the add-in. Update the closed file with the command above; the next reload launcher opens the new version. Do not replace an add-in that is executing its own update. Existing hot reload through WorkbookUpdater retains its VBA lifecycle coordinator; it does not launch this installation script.

Existing workbooks are backed up, including unsaved edits, in `.backup/modules-<id>` beside the workbook before cleanup. Hot reload uses the same `.backup` root with `reload-<timestamp>` operation directories. An import or initialization failure does not save the partial result. Loaded workbooks can remain partially modified in memory: recover from the reported backup rather than saving them. There is no automatic rollback or full-project compile assertion. The original Excel event, screen-update and calculation settings are restored.

Run `Test-WorkbookModules.ps1` for isolated Excel checks of Unicode, class headers, native imports, schema masks, document routing, obsolete-module removal, backups and preflight rejection. Run `Test-Updater.ps1` against a newly built add-in to verify the hot-reload lifecycle. Tests do not modify user workbooks.
## Контракт кодировки

`modules.json` и исходники `.vba` читаются как строгий UTF-8 (BOM допустим). Некорректные последовательности и символ замены U+FFFD прекращают операцию до изменения VBA-проекта. Строковые литералы преобразуются в выражения `ChrW$`, включая UTF-16 surrogate pairs. Comments are passed through unchanged; new and modified code comments must be in English. Идентификаторы VBA должны быть ASCII. Unicode-литералы в `Const` нужно заменить инициализацией переменной во время выполнения: вызов `ChrW$` не является константным выражением.

Нативные текстовые экспорты `.bas`, `.cls`, `.frm` принимаются только в ASCII, без BOM. Это исключает зависимость текстового импорта VBE от системной ANSI-кодировки. Для Unicode-кода используйте `.vba`; подписи элементов UserForm задавайте из Unicode-кода после создания формы. Бинарные `.frx` сохраняются без перекодирования; содержимое стороннего бинарного ресурса не проверяется на переносимость текста.

Сообщения Excel и `Err.Description` передаются как Unicode-строки напрямую, без промежуточной ANSI-конвертации. Диагностические логи VBA уже используют UTF-16 (`TristateTrue`), а не UTF-8: читать их нужно как UTF-16. Кодировку существующих логов не меняем при дописывании. Повреждённые ранее символы `?` автоматически не восстанавливаются.

Проверка: `Test-SourceEncoding.ps1` проверяет оба импортёра, некорректный UTF-8, Unicode-литералы, статус-бар и запись/чтение UTF-16 лога. Проверки `Test-WorkbookModules.ps1` и `Test-Updater.ps1` проверяют импорт и жизненный цикл.