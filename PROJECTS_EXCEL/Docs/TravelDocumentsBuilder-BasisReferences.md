# Travel document basis references

## Runtime contract

The institution workbook owns the reusable institution names, grammatical forms, officials, document types and compatibility rules. `wsConfig` owns source descriptors, column mappings and application messages. XAML owns the form layout and styles.

The editor first selects a source category, then an institution from the candidate table. The selected record supplies its actual group. A numbered military hospital therefore uses Hospital rules even when found through the military-unit source. Command institutions use separate ChiefCommand, GeneralStaff and LandCommand groups to prevent incompatible titles.

Reference tables:

| Table | Contract |
| --- | --- |
| tbBasisOfficials | Stable Id, nominative display name and text used in a basis |
| tbBasisOfficialRules (Посади установ) | Allowed institution group and official Id pairs |
| tbBasisDocumentTypes | Stable Id, display name, genitive form, template, required fields, visible fields and available document forms |
| tbBasisDocumentRules (Документи установ) | Allowed institution group and document Id pairs |
| tbCommands | Group, unique name used as the lookup key, and genitive name of a military authority |
| tbCourts | Group, unique name used as the lookup key, and genitive name of a court |
| tbBasisFields | Public field code, readable name, description, data type and input/computed usage |
| tbBasisFragments | Public fragment code, readable explanation, source-field condition and output template |
| tbBasisContract | Contract version, required capabilities and a human-readable syntax legend |
| tbEnumInstitutionGroups (Переліки) | Editable list of permitted institution groups used by Excel data validation |

Keep Id and display names unique in each definition table. Do not change existing Id values. Rules refer to Id, not the display text. The original institution tables and their declensions retain their existing schema.

Command and court lookup keys use `Назва`; their separate `Id` columns are no longer required. Names must be unique within each of those institution tables. Existing history keeps its saved keys and text; newly selected institutions capture their name as the key. Lookup queries select a physical source column once even when both Key and Name map to it. Definition Ids for document types and officials remain unchanged.

## Document fields

`Поля` is a pipe-separated list of public input-field codes. Other editor fields are hidden and cleared when the document type changes. `Обов’язкові поля` uses `&` for AND, `/` for OR and explicit parentheses for grouping. For example, `((document.number&document.date)/(registration.number&registration.date))` requires either the document's own registration pair or an incoming registration pair. A report requires `(document.author&((document.number&document.date)/(registration.number&registration.date)))`. Mixed operators at the same nesting level are rejected rather than assigned precedence. Partially filled registration pairs are still rejected independently.

`|` is also accepted as AND for compatibility. Saved history snapshots without parentheses retain the original separator-based interpretation. New drafts use the parenthesis-aware parser and never rely on legacy precedence. The parser reads every branch, rejects unknown fields, empty groups and unbalanced parentheses, and limits nesting to 64 groups.

Templates reference only public field or fragment codes. A field such as `{document.number}` inserts its raw value; the fragment `{document.numberText}` includes the optional number prefix. Fields have `text`, `date` or `choice` types and `input` or `computed` usage. The field dictionary provides the captions used by the dynamic TDB editor.

Fragments use `nonempty`, `equals`, `except` or `default` conditions. `except` requires a nonempty source value different from the configured value. A default applies only when no conditional case matches; if no case or default matches, the fragment is empty. Conditional cases for one code share one source field and must be mutually exclusive. Fragments may reference other fragments; cycles are rejected. Wording, including prefixes and optional grammatical phrases, is owned by these tables.

`basisChain.contract.bindings` is the TDB adapter: for example, `document.number=Form.BasisNumber`. `basisChain.contract.bindingTypes` declares the types of the local adapter targets, so a choice cannot be mapped to a plain text or date target accidentally. Those internal paths never appear in the shared field dictionary or document profiles. `basisChain.contract.legacyTokens` is used only to migrate previously authored profiles. Other consumers provide their own mappings and supported capabilities. The implementation also pins the version and feature set it actually supports; changing a configuration declaration does not add engine capabilities.

The dependency loader validates the staged contract before publication: supported version and capabilities, unique codes, all mappings and data types, input-field usage, visible versus required fields, template tokens, fragment cases and cycles. Unknown tokens are errors, not literal text or silent empty values. Existing history retains its compiled internal field names; newly captured templates compile public fragments into frozen `{Value.<internal field>}` substitutions and continue to store final rendered text. No history schema change or revision rewrite is needed.

The acting-duty selector is independent of the official. Copy/extract selection switches the document type to its stored genitive form. Incoming registration is separate from the original document number and date. A report author and the issuer of a referenced personnel order are entered as complete grammatical phrases in the genitive; they are not automatically declined from a person's name.

The personnel-notification profile supports one referenced order with issuer, number, date and optional scope. It does not create an unrestricted recursive document graph. A previously issued travel certificate can be added as a basis; the newly generated certificate is not added automatically.

Each added basis stores its institution key/group, definition Ids, grammatical text, template, required-field specification and final rendered text. Reading history validates copies instead of rewriting saved rows. Existing records without the new snapshots retain the legacy formatting path. Editing and adding a basis captures the current definitions again. Each new field is also included in revision snapshots.

Indefinite business trips keep the existing duration mechanism. Notes such as arrival, unauthorized absence or exclusion do not create a different trip kind in this implementation.

## Workbook upgrade

`Build-TDBBasisReferences.mjs` creates an artifact-tool seed workbook. Run it with the bundled Node runtime and `@oai/artifact-tool` available through a temporary directory's `node_modules` junction. Pass the output directory as its only argument.

`Upgrade-TDBBasisWorkbooks.py` merges those authored tables into the institution workbook without re-exporting unrelated sheets. Existing definitions and custom rows are retained; only missing Ids and rule pairs are appended. Original institution and place worksheets remain byte-for-byte intact. New reference-sheet styles are imported with remapped style indexes.

Example arguments for the bundled Python runtime:

```text
Upgrade-TDBBasisWorkbooks.py --references <seed.xlsx> --institutions <Установи.xlsx> --history <TicketsHistory.xlsx> --output-dir <upgrade-directory>
```

Without `--apply`, only upgraded copies are written. With `--apply`, timestamped backups are created before replacing the source files. Close the source workbooks before applying, so a later save from an old Excel session cannot overwrite the upgrade. The utility does not import or execute VBA.

`--explicit-required-only` updates only the required-field expressions in the institution workbook, using their previous grouping to insert parentheses. This mode does not need a seed workbook and never modifies history snapshots. Existing parenthesized expressions are preserved.

`--rename-rule-sheets-only` renames the legacy rule tabs and their titles to `Посади установ` and `Документи установ`. Stable table names and rule data do not change. This focused mode does not need a seed workbook and does not touch history. Full upgrades also recognize and rename legacy tabs before merging their data.

`--group-enums-only` adds the final `Переліки` sheet and the `tbEnumInstitutionGroups` table, then installs stop-style Excel list validation in every other `Група установи` column. The validation uses `INDIRECT("tbEnumInstitutionGroups[Група установи]")` and covers the data column down to the last worksheet row so new rows are covered. Existing unrelated validations remain intact. The enumeration table itself stays editable and is not validated against itself. Existing nonempty group values must already belong to the enumeration; the upgrade reports unknown values instead of silently admitting them. Full upgrades also refresh this validation.

The history upgrade appends missing configured columns to both tables. It preserves all existing rows, revisions and retired column values. Runtime requires every configured column but tolerates surplus columns. `FoodDepartureDate` remains unsupported even if a legacy workbook still contains it.

After applying the workbook upgrade, reload the TDB configuration and modules, then reload dependencies. A stale configuration cannot address the new tables and fields.

## Validation

Run `Test-TDBBasisInstitutions.py` and `Test-TDBBasisDocuments.py` with the bundled Python runtime. The second test accepts explicit upgraded and original workbook paths to check preservation before applying. It checks reference integrity, incompatibilities, sample document phrases, history schemas, frozen-text source contracts and XAML bindings. These are static/data checks; Excel interaction and Word export still require runtime verification.