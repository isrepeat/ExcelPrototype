# PeopleRegistry

`PeopleRegistry.xlsm` contains the XAML interface and `wsConfig` / `tbConfig`. The active profile is `PeopleRegistry`; configuration is loaded from `config/PeopleRegistry/wsConfig.txt` and markup from `ui/PeopleRegistry/MainPage.xaml`.

`PeopleRegistryBase.xlsx` contains the registry and origin aliases. It is stored beside the application workbook and selected by `PeopleRegistry::registry.path`.

The active VBA sources are `ex_PRM.vba`, `obj_PRM_PgMain.cls.vba` and `ThisWorkbook.vba`, plus the common runtime selected by `vba/modules.json`. Registration and replacement in the external database remain unconnected.

`legacy/` preserves the former combined-workbook implementation and its maintenance scripts for reference. It is excluded from the active module profile. Do not run its scripts against the new application or database. The original combined workbook retains its original configuration.