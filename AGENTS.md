# Agent Instructions

Write all new or modified code comments and developer-facing documentation in English. Do not add Cyrillic comments. Keep technical identifiers, API names, and configuration keys in English. Store user-facing messages, labels, captions, table headers and status templates in wsConfig; do not hardcode them as VBA or markup string literals. Unicode is allowed in configuration values and in source identifiers such as workbook paths, worksheet names and source column headers. Existing comments can be translated in a separate change.

Wrap public procedures in standard VBA modules in a comment-delimited namespace block, following the established format: `namespace API {` and `} // namespace API`. Group public procedures by a more specific namespace when the module already uses one. This rule does not apply to class modules or interface class modules; use the class section format below and do not leave namespace opening or closing comments in them.

## VBA class structure and method style

- Флаги жизненного цикла объявляются первыми полями конкретного класса: сначала `Private m_isInitialized As Boolean`, затем `Private m_isDisposed As Boolean`. После них оставляется ровно одна пустая строка перед остальными полями.

- Match the method and comment style of neighboring project classes. Use three-line section comments: `' //`, `' // Lifecycle` (or the appropriate section name), `' //`.
- Every concrete class must contain a single Lifecycle section containing only the lifecycle callbacks `Private Sub Class_Initialize()` and `Private Sub Class_Terminate()`. Public `Initialize` and `Dispose` methods must be the first two procedures in the API section, in that order. Preserve required initialization arguments. Interface classes contain only the necessary public contract and have no Class_Initialize/Class_Terminate callbacks. Do not add Initialize/Dispose to an interface unless lifecycle management through that interface is an explicit requirement; normally the owner manages the concrete class lifecycle.
- Class_Initialize creates only minimal internal state, such as collections and initial scalar values. It must not call the public Initialize method or open external resources. An empty Class_Initialize is valid when no internal setup is needed. Initialize explicitly prepares the object with its data/dependencies and returns a Boolean that callers must check. Dispose must be safe to call repeatedly; Class_Terminate calls Me.Dispose as final cleanup. Resource services/executors must reject use before successful explicit initialization.
- Use two separate Boolean members in concrete classes: m_isInitialized and m_isDisposed. Both start False. Only successful initialization sets m_isInitialized=True. Apply these flags to every concrete PROJECTS_EXCEL class, including existing UI classes. Preserve existing interface signatures: where Configure is the initialization entry point, track successful Configure with the same flags; existing Sub Initialize contracts may remain Subs. Dispose sets m_isInitialized=False and m_isDisposed=True and is idempotent. Initialize must reject already initialized or disposed instances without changing their state; it must not call Dispose as an initial reset. Create a new instance for another lifecycle. Cleanup after initialization failure may call Dispose to release partially acquired resources. Interfaces do not contain these flags.
- Keep all Property Get/Let/Set procedures together in a Properties section immediately after Lifecycle and before Interface, API, or Private sections. Do not scatter properties among API methods. Omit the Properties section when there are no property procedures.
- Qualify calls to a class's own public methods with `Me`, including lifecycle callbacks, error handlers, cleanup paths and interface forwarding methods; for example, `Me.Dispose` and `Me.TryExecute(...)`. Call Private helpers directly, without `Me`. Do not qualify a Function or Property return-value assignment with `Me`.
- Keep internal helpers Private and retain the established `private_` name prefix. Do not change them to Friend or Public solely to qualify calls through `Me`; this caused the observed VBA member-resolution error.
- Put each parameter of a method with multiple parameters on its own continuation line, using the neighboring project format. Keep individual variable declarations on separate lines and a blank line between declarations and executable statements.
- Use ordinary multiline If/For/Select Case blocks. Do not compress multiple statements with colons or place loop bodies and Next on one line.

When the user requests an upload of changes, create one or more logical local commits from the current changes. Do not run `git push` unless the user explicitly requests it.

Write every created commit message in English.

## End of file

- The final byte of every text source, configuration and documentation file must belong to the last meaningful line. Do not leave a trailing newline, empty line or trailing whitespace.
- After edits, check every existing text file in the union of `git diff --name-only` and `git ls-files --others --exclude-standard`. Compare its final byte numerically against 9, 10, 13 and 32; none is allowed.

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

## Configuration and user-facing text in wsConfig

- Store every application configuration in wsConfig / tbConfig, including lookup sources, schemas, limits, search rules, display projections and field mappings. Do not introduce separate XML, JSON or other configuration files for these settings. XAML describes UI structure and bindings; runtime settings come from wsConfig. Existing tool manifests are outside this application-configuration rule.
- Keep all user-facing messages and UI text in the mode's `wsConfig.txt`, loaded into the workbook's `wsConfig` / `tbConfig`. Read them through the existing configuration API and expose UI values through bindings. This applies to text in any language, not only Cyrillic.
- Use configuration templates with named placeholders such as `{count}` and `{minChars}` for dynamic messages. Keep only substitution logic in VBA.
- Use actual tab characters between the three fields of each wsConfig text row. Never write the two literal characters `\t` as a separator. Validate the tab-separated structure when editing configuration files.
- Do not silently substitute hardcoded text for a missing required configuration key. Report the missing key specifically and stop the operation. A minimal bootstrap diagnostic identifying the unavailable configuration key is permitted when the configured message cannot be loaded.
- Source workbook paths, worksheet names and column headers describe external data and may contain Cyrillic; they are not localized UI captions. Keep display captions separately configurable.

## XML and XAML formatting

- Order XML/XAML attributes as follows: `name`, `type`, `dataContext`, `itemsSource`, `row`, `column`, `rowSpan`, `columnSpan`, `orientation`, `style`, `value`, `label`, `caption`, `command`. Within the paired groups, row precedes column, rowSpan precedes columnSpan, and label precedes caption. Omit absent attributes without blank lines; the next present attribute takes their place. Attributes outside this list follow the listed attributes, retaining their relative order. Namespace declarations remain before ordinary attributes on the root element.
- Put each attribute on its own line. The first attribute may follow the tag name on the opening line; align subsequent attributes with it. Never put several attributes on one line.
- Indent nested elements consistently with the surrounding markup. Keep closing tags aligned with their opening elements.
- Do not use empty `grid` elements as spacers. Use `stackPanel` with an explicit `orientation` for the existing empty-container spacing pattern.
- Validate XML syntax and nesting after editing markup, without launching Excel or executing VBA.

## Visible file edits

- Make manual file changes through `apply_patch` or an equivalent tool that displays the changed files and patch in the chat. Do not hide manual edits in shell commands or scripts that rewrite files. For generated files or required byte-level normalization, identify the affected files and purpose in the chat.

## Validation without VBA execution

- Full page rendering and layout rebuilds must disable `Application.ScreenUpdating` before clearing or changing the page and restore its previous value on every exit, including errors. Keep this guard in the rendering pipeline; ordinary cell selection must not toggle screen updating merely because focus changed.

- Do not automatically compile the VBA project or run VBA macros/tests as part of source edits. Compilation dialogs interrupt the user. Use static source, XML, configuration, naming and final-byte checks instead.
- Do not import changed sources into or update the working workbook unless the user explicitly requests that action. Do not treat a general implementation request as permission to start VBA compilation or execution.
- Run VBA compilation or macro-based validation only when the user explicitly authorizes it. Report when runtime behavior remains unverified; static checks do not establish VBA runtime correctness.