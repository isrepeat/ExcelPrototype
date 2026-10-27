# PeopleRegistry

`PeopleRegistry.xlsm` contains the XAML interface and `wsConfig` / `tbConfig`. The active profile is `PeopleRegistry`; configuration is loaded from `config/PeopleRegistry/wsConfig.txt` and markup from `ui/PeopleRegistry/MainPage.xaml`.

`PeopleRegistryBase.xlsx` contains the registry and origin aliases. It is stored beside the application workbook and selected by `PeopleRegistry::registry.path`.

The active VBA sources are `ex_PRM.vba`, `obj_PRM_PgMain.cls.vba`, `ex_PR_Registry.vba` and `ThisWorkbook.vba`, plus the common runtime selected by `vba/modules.json`.

Open the configured `PeopleRegistryBase.xlsx` in the same Excel instance before registering or replacing an ID. The application resolves the full configured path, never opens or closes the database itself, and rejects read-only workbooks, protected history sheets and invalid table schemas. It reads the currently open history and aliases, including unsaved edits. Successful operations modify only the open database in memory. The user saves it explicitly; the application never saves either workbook or reads or changes RUH.

The registration and replacement buttons use XAML commands and binding values instead of fixed cell addresses. ID generation, collision checks, same-name confirmation and flattened replacement links use only the database. The last successfully registered input is remembered within the session so a repeated click resolves the same ID. Replacement requires explicit confirmation. A write failure restores the previous history table in memory and clears the successful result. Closing the database without saving discards changes.

Ordinary registration appends one row without rewriting existing records. Replacement updates the history links. Diagnostic stages `PR.Registry.ReadHistory`, `ReadAliases`, `ProcessAndWrite`, `WriteHistory` and `Total` separate database work from the shared UI timings. The database can remain closed while the UI opens or refreshes; it is required when a mutation command is executed. Identifiers that exist only in other workbooks must first be registered in the database.

`legacy/` preserves the former combined-workbook implementation and its maintenance scripts for reference. It is excluded from the active module profile. Do not run its scripts against the new application or database. The original combined workbook retains its original configuration.