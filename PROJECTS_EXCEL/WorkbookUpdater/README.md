# Workbook updater

`WorkbookUpdater.xlam` owns the deferred import operation and the retained runtime context. The target workbook remains open. PERSONAL.XLSB only captures the active workbook, loads the updater, and submits the request. Normal reloads no longer launch PowerShell or close the workbook.

## Installation

1. Build the add-in with `powershell.exe -NoProfile -ExecutionPolicy Bypass -File PROJECTS_EXCEL\WorkbookUpdater\Build-Updater.ps1`. This wrapper uses the common `Scripts/PowerShell/Update-WorkbookModules.ps1` importer and the local `modules.json` profile. Excel and trusted access to the VBA project object model are required. The script creates its own hidden Excel instance and refuses to overwrite an existing output. The generated add-in is ignored by Git. To update an existing unloaded add-in, use the common importer's `-WorkbookPath`, `-VbaFolderPath` and `-Profile WorkbookUpdater` parameters; see `Scripts/PowerShell/WorkbookModules.README.md`.
2. For a saved workbook with no VBA code, Ctrl+Alt+R performs the initial installation from its configured source profile. No existing lifecycle module is required. All components must have zero code lines, no forms or unsupported designers may exist, and a failed-update marker must be absent. A workbook that already contains code without the lifecycle contract is deliberately rejected; use the common PowerShell importer for its explicit migration.
3. Update PERSONAL.XLSB from `MacrosExcel\PERSONAL` using the existing personal-workbook installer. Ctrl+Alt+R keeps its existing public handler names.
4. Activate the target workbook and press Ctrl+Alt+R. Its `tbConfig` must contain the existing profile ID, relative VBA folder and relative UI folder. The launcher locates the add-in at `<parent of vba folder>\WorkbookUpdater\WorkbookUpdater.xlam`.

The current source changes and generated add-in do not automatically modify an already open target workbook or PERSONAL.XLSB. Updating the source file is different from importing it into a running workbook.

## Reload protocol

`Running -> Requested -> Prepared -> Importing -> Initializing -> Running`

Any failure after the request enters `Faulted`. There is no automatic retry or return to `Running` after a partial import.

- Preflight expands `modules.json`, rejects duplicate component names and missing document modules, and reads the source plan before requesting shutdown. UTF-8 `.vba` code is retained in memory. Native `.bas`, `.cls` and `.frm` exports are staged; referenced `.frx` resources must exist.
- `StopRequested = True` rejects new supported entry calls. The captured workbook object determines the target even if another workbook becomes active.
- Initial installation creates its context in the add-in, rechecks that the target is still empty immediately before changes, and skips the old runtime's prepare call. It uses the same backup, import, initialization and failure blocking as regular reload. A previously initialized runtime whose code was manually removed cannot use this shortcut.
- The add-in owns the exact OnTime timestamp and procedure name. It waits for `ActiveCalls = 0`, with a 30-second deadline. Closing the target cancels the pending callback.
- `SaveCopyAs` writes a backup, including current unsaved edits, before disposal and code changes. The backup and staged exports are retained in `.backup/reload-<timestamp>` beside the target file. The PowerShell installer uses the same `.backup` root with separate `modules-<id>` operation directories.
- Strict hotkey detachment and diagnostic flushing precede disposal of page, binding, UI, style and other current runtime roots. An unsuccessful prepare result prevents import.
- Import replaces standard/class/form components and updates existing workbook/worksheet components by their CodeNames. Removed profile event handlers are cleared.
- The shared dictionary is retained in the add-in registry and attached to the new target runtime. This handles target module-variable resets caused by recompilation.
- After importing, the add-in returns from its callback with the runtime still blocked. Initialization runs in a separate OnTime callback after Excel's deferred project reset. Otherwise a rendered UI can survive while its module-level command registry disappears. Cancellation is permitted only before importing starts.
- Ctrl+Alt+D submits `fn_RequestClear` through the same coordinator. It blocks entry, waits for active calls, backs up the workbook and completes lifecycle cleanup before removing any component. It clears document code, saves the empty project and releases the retained context. No target methods are called after deletion, and Ctrl+Alt+R can perform initial installation again. Populated legacy projects without lifecycle must use explicit migration or recovery.
- Initialization must report success. Only then is the entry gate reopened. The updater restores the original Excel event, calculation and screen-update settings on both success and failure. It does not save the updated target workbook automatically.
- A workbook name `_RuntimeReloadBlocked` preserves the blocked state if target VBA variables are reset. The marker is created after the clean backup and removed only after successful initialization. Context entries are released when the target closes.

`VBProject.Mode = Design` is not used as an in-callback prerequisite: Excel returns `Run` while the add-in's VBA callback executes, even when the target has no active entry calls. The stop protocol and entry counter provide the relevant check.

## Contracts

Current modules already expose `fn_Module_Dispose`; owned classes expose `Dispose` or implement the existing page/control disposal interface. Only resource-owning modules need real cleanup. Disposal must be idempotent, release child objects and break reference cycles; failure must propagate to the lifecycle coordinator. `Class_Terminate` is a fallback, not the shutdown protocol.

The context fields are `StopRequested`, `ActiveCalls`, `Phase`, `Generation`, and an error message on failure. Keep this object independent of replaceable VBA class types. Do not mutate these fields from business code.

Supported event handlers, the shape bridge and the diagnostic hotkey enter and leave the runtime gate. Internal methods do not all need context parameters. A new external entry point must use `fn_TryEnter` and always balance an accepted entry with `fn_Leave`, including error exits. A long operation that yields through `DoEvents` can accept this context and check `fn_StopRequested` between steps. Cancellation is cooperative; it does not undo completed writes.

Future OnTime tasks, COM event subscriptions, API timers, forms and external object owners must be added to the prepare/dispose protocol before they are supported for reload. The current target sources contain no independent OnTime tasks or Windows callbacks. Arbitrary direct calls that bypass the entry gate, external references to old class instances and manual project edits are outside the guarantee.

The lifecycle initializer invokes the callback specified by the required wsConfig key `ThisWorkbook::initializer`. The common page factory uses `ThisWorkbook::pageFactory`. Both values must name a public function as `Module.Procedure` in the target workbook; missing or invalid values stop the operation. Each mode owns its initializer, concrete page creation and resource ordering. Profiles import only common sources and their own mode sources.

## Failure recovery and validation

After a failed reload, keep the runtime blocked. Close without saving and use the reported backup if code changed. The updater does not undo changes that successful or failed initialization already made to sheets; the backup is the recovery boundary.

An imported component's name and type are checked, and runtime initialization is verified. Import success does not prove every unreachable VBA procedure compiles. Full `Debug -> Compile VBAProject` and application regression checks remain necessary before releasing a source version. The updater does not use brittle VBE menu automation to claim a full compilation check.

Run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File PROJECTS_EXCEL\WorkbookUpdater\Test-Updater.ps1` for an isolated Excel integration test. It creates temporary fixture workbooks, uses the real lifecycle module, and verifies entry blocking, waiting for active calls, a captured source snapshot and target, retained context, backup creation, exact callback cancellation and initialization failure. It does not touch user workbooks. Fixture artifacts are preserved for inspection.

`Test-ProjectSources.ps1 -WorkbookPath <path to the existing PersonalEventBuilder workbook>` imports the real project sources into a temporary copy with workbook events disabled, then checks runtime entry accounting and the shutdown contract. It does not replace or save the original workbook. This is a contract smoke check, not a full application regression test.