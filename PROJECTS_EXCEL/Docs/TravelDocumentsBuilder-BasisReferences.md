# Travel document basis architecture

## Responsibilities

- The institution workbook contains source information only: institutions, groups, names and declensions.
- wsConfig owns source/column descriptors, technical identifiers, field definitions, document profiles, allowed relationships, validation rules and choice-value mappings.
- BasisTemplates XML owns wording, punctuation, grammatical choices and optional text blocks.
- VBA validates the data and executes the text templates. The resulting string is shared by the form preview and Word export.

There is no workbook contract or capability record. Configuration and XML are validated by the consuming code before publishing a dependency generation. No historical ticket or revision columns are removed.

## Reference workbook

| Table | Purpose |
| --- | --- |
| tbBasisOfficials | Nominative name and genitive text |
| tbBasisDocumentTypes | Document name and genitive form |
| tbCommands, tbCourts | Institution group, unique name used as key, genitive name |
| tbEnumInstitutionGroups (Переліки) | Allowed institution groups for native Excel validation |

Military units use their unit designation as the key; medical facilities use their designation. Courts, authorities, officials and document names use unique names as source keys. Stable technical document and official identifiers exist only in configuration and historical snapshots.

Configured profile `Fields` lists input-field codes separated by pipes. `Required` expressions use AND (`&`), OR (`/`) and explicit parentheses: `((document.number&document.date)/(registration.number&registration.date))`. Mixed operators at one nesting level are rejected. Registration pairs are also checked for partial entry. Legacy snapshots retain the original separator interpretation.

The selected institution's actual group controls available documents. Officials follow that group unless `basisChain.model.officialGroupOverrides` overrides it for the selected institution kind. The configured `Unit~Unit` override gives every numbered military unit the military-unit official list, including medical units whose source group is Hospital. Selecting a hospital keeps its hospital official list. Both selection and save validation apply the same override. Search projects each physical column once when Key and Name share a header.

## Configuration adapter

`basisChain.model.bindings` maps template data to local form/storage fields, for example `document.number=Form.BasisNumber`. `bindingTypes` declares the target types. `model.Fields.<code>` supplies the caption, description, type and usage.

`model.Documents.<id>.Name` and `model.Officials.<id>.Name` select a unique source row by its name; the remaining grammatical text comes from that row. Document profile `Fields`, `Required` and `Forms` live in this configuration branch. `model.DocumentRules.<group>` and `model.OfficialRules.<group>` list permitted configuration identifiers. Every branch has an explicit `keys` list; unknown references, missing source names and duplicate source matches stop loading. Renaming a configured source key requires updating its configured `Name`.

`ex_TDB_BasisConfig` combines source names/declensions with those settings in memory. Neither a technical-field dictionary nor a contract is loaded from the source workbook.

`basisChain.textTemplates.profiles.<document Id>` selects a root XML file. Several profiles share Common.xml. `folder` is resolved relative to the TDB workbook directory; `ownOrder` selects the automatic first basis document.

`valueMaps` converts UI captions to stable template values. The current form still displays localized choices. XML receives `original/copy/extract` for document.form and `true/false` for issuer.acting. The "no scope" UI choice becomes an empty scope. Closed mappings reject unknown UI values; open mappings preserve unmatched text. Renaming a choice caption requires updating its adapter mapping, not XML conditions.

## XML language

Each file has a `<textTemplate version="1">` root. Formatting whitespace between elements is ignored. Literal prose must be inside `<text>`. Nonempty rendered parts are joined with one space; final CR/LF/tab and repeated ASCII spaces are normalized. Field values are inserted as data and are not parsed as templates.

| Element | Meaning |
| --- | --- |
| text | Literal wording |
| value field="code" | Raw public value |
| ifFilled field="code" | Include children when the value is nonempty |
| if field="code" | Include children for the normalized Boolean value true |
| choose / case value / otherwise | Exact choice matching, then optional default |
| group prefix / suffix | Attach punctuation to nonempty rendered children |
| include file="Name.xml" | Reuse another file in the same template folder |

DocumentTitle.xml chooses nominative or genitive names for original/copy/extract. Requisites.xml and Registration.xml supply optional labels and punctuation. Issuer.xml supplies the optional acting-duty phrase. Root profiles arrange these reusable blocks.

Only declared fields, elements and attributes are allowed. Choice cases must belong to the declared closed value map where available. Missing files, malformed XML, DTDs, external entities, path traversal, include cycles, duplicate cases and unsupported versions are rejected. Templates are cached as part of the staged dependency generation; file stamps are verified before publication.

The automatic own-order template has a separate context: order.number, order.date and ownUnit.genitive. Pending number/date placeholders are supplied as data for the initial form preview.

## History and Word

Each captured document stores its structured fields, required specification, definition Ids and final BasisRenderedText. BasisTemplate is marked `xml.v1` for new documents; old compiled templates and legacy snapshots remain readable. A marked XML snapshot without rendered text is invalid.

Saved events are rendered from approved text, not current XML assets. The automatic first document is recovered by removing the exact saved child-document suffix from the approved basis string. It is held in Form.BasisOwnText for preview and subsequent edits. Clearing the event clears that state. No history schema migration is needed.

Word uses the approved root Basis value. Historical consistency checks compare saved child text to the saved root rather than regenerating text through today's templates. New or explicitly edited documents use the current loaded XML generation.

The report author and referenced order issuer are still entered as complete genitive phrases; automatic personal-name declension is outside this change. Indefinite trips retain their existing duration mechanism.

## Upgrade and verification

`Build-TDBBasisReferences.mjs` authors information-only seed reference tables using the bundled artifact library; execute from the repository root. It does not emit technical fields, rules, profiles or contract metadata.

`SourceOnly-TDBInstitutions.py --institutions <book> --output-dir <directory>` writes an information-only copy. Add `--apply` to create a timestamped backup and replace the source. Close the workbook first to avoid a later save from an old session overwriting the upgrade. The migration checks that settings match configuration before deleting technical columns and the rule/field/contract sheets. It preserves the retained source values, styles and group validation and does not touch TicketsHistory.

The legacy `--xml-basis-only` switch delegates to the information-only migration. The group enumeration uses dynamic structured-reference list validation with stop errors; unrelated validations are preserved.

Run Test-TDBBasisInstitutions.py, Test-TDBBasisDocuments.py, Test-TDBBasisContract.py and Test-TDBGroupEnums.py with the bundled Python runtime. XML tests exercise actual assets with an independent fixture reader, all profile families and malformed inputs. They do not execute or compile VBA.

After upgrading, reload TDB configuration/modules and then dependencies. Excel interaction and Word export require separate runtime verification.