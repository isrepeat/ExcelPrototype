# Agent Instructions

Write all new or modified code comments and developer-facing documentation in English. Do not add Cyrillic comments. Keep technical identifiers, API names, and configuration keys in English. Unicode user-facing string literals are allowed. Existing comments can be translated in a separate change.

Wrap public procedures in standard VBA modules in a comment-delimited namespace block, following the established format: `namespace API {` and `} // namespace API`. Group public procedures by a more specific namespace when the module already uses one. This rule does not apply to class modules or interface class modules; use the class section format below and do not leave namespace opening or closing comments in them.

## VBA class structure and method style

- Флаги жизненного цикла объявляются первыми полями конкретного класса: сначала `Private m_isInitialized As Boolean`, затем `Private m_isDisposed As Boolean`. После них оставляется ровно одна пустая строка перед остальными полями.

- Match the method and comment style of neighboring project classes. Use three-line section comments: `' //`, `' // Lifecycle` (or the appropriate section name), `' //`.
- Every concrete class must contain a single Lifecycle section with `Private Sub Class_Initialize()` and `Private Sub Class_Terminate()`, and public `Initialize` and `Dispose` methods. Preserve required initialization arguments. Interface classes contain only the necessary public contract and have no Class_Initialize/Class_Terminate callbacks. Do not add Initialize/Dispose to an interface unless lifecycle management through that interface is an explicit requirement; normally the owner manages the concrete class lifecycle.
- Class_Initialize creates only minimal internal state, such as collections and initial scalar values. It must not call the public Initialize method or open external resources. An empty Class_Initialize is valid when no internal setup is needed. Initialize explicitly prepares the object with its data/dependencies and returns a Boolean that callers must check. Dispose must be safe to call repeatedly; Class_Terminate calls Me.Dispose as final cleanup. Resource services/executors must reject use before successful explicit initialization.
- Use two separate Boolean members in concrete classes: m_isInitialized and m_isDisposed. Both start False. Only successful initialization sets m_isInitialized=True. Apply these flags to every concrete PROJECTS_EXCEL class, including existing UI classes. Preserve existing interface signatures: where Configure is the initialization entry point, track successful Configure with the same flags; existing Sub Initialize contracts may remain Subs. Dispose sets m_isInitialized=False and m_isDisposed=True and is idempotent. Initialize must reject already initialized or disposed instances without changing their state; it must not call Dispose as an initial reset. Create a new instance for another lifecycle. Cleanup after initialization failure may call Dispose to release partially acquired resources. Interfaces do not contain these flags.
- Keep all Property Get/Let/Set procedures together in a Properties section immediately after Lifecycle and before Interface, API, or Private sections. Do not scatter properties among API methods. Omit the Properties section when there are no property procedures.
- Qualify calls to a class's own methods with `Me`, including calls from lifecycle callbacks, error handlers, cleanup paths, interface forwarding methods, and other private/public methods. For example, use `Me.Dispose`, `Me.TryExecute(...)`, and `Me.private_TryRead(...)`. Do not qualify a Function or Property return-value assignment with `Me`.
- VBA does not expose Private helper methods through Me. When a helper must be called through Me, declare it Friend while retaining the established private_ name prefix; do not make it Public solely for qualification. Verify these calls by compiling the VBA project.
- Put each parameter of a method with multiple parameters on its own continuation line, using the neighboring project format. Keep individual variable declarations on separate lines and a blank line between declarations and executable statements.
- Use ordinary multiline If/For/Select Case blocks. Do not compress multiple statements with colons or place loop bodies and Next on one line.

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