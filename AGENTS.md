# Agent Instructions

Write all new or modified code comments and developer-facing documentation in English. Do not add Cyrillic text to source code, comments, or configuration keys. Keep technical identifiers, API names, and DSL keywords in English.

Wrap every module's public VBA procedures in a comment-delimited namespace block, following the established format: `namespace API {` and `} // namespace API`. Group public procedures by a more specific namespace when the module already uses one.

When the user requests an upload of changes, create one or more logical local commits from the current changes. Do not run `git push` unless the user explicitly requests it.

Write every created commit message in English.

## End of file

- Do not add a trailing newline after the last character when editing files. Remove it when present.

## Error handling

- Do not use silent default fallbacks, such as automatically selecting `Default`.
- When required data, a profile, or a file is missing, show a `MsgBox` with a specific description and stop the current operation.

## Naming constraints in VBA and Excel

- Account for VBA and Excel name-length limits when adding VBA code or declarative UI controls.
- An Excel `Shape` name, including automatically added prefixes such as `btn_`, must not exceed 31 characters. Otherwise, Excel truncates the name and the registered runtime route may not match the actual `Shape` name.
- Before completing changes, verify final Excel-object names together with all generated prefixes and suffixes.

## Custom class variable naming

- When practical, name a variable after its assigned custom class. For example, name an `obj_PrsnlEvntBuilderCfgParser` instance `prsnlEvntBuilderCfgParser`.
- This is a preferred style rather than an absolute requirement. Apply it first in simple methods without several similarly named dependencies.
- Semantic prefixes and suffixes are allowed, but the variable-name root should remain recognizably related to the class name when possible.

## Mode-specific VBA logic

- `PrototypeNew` is a universal engine and must not directly create mode-specific classes by default.
- If a mode requires specific VBA logic, place its class in `PrototypeNew/vba/[5] pages/<ModeName>/`.
- Explicitly specify a mode-specific class name in the mode XML configuration. Treat a missing required class name as a configuration error, show a specific `MsgBox`, and stop the operation.
- Shared code may contain a universal contract, lifecycle framework, and factory entry point. Select and create a concrete implementation only by the class name read from configuration.
- Do not add checks for a specific `ModeName` or mode-specific processing rules to universal controller/parser classes when that logic can be isolated behind a shared interface.