# History-to-Word performance

Each export rebuilds the complete selected order from the empty template. After insertion and verification, all Begin/End export markers, including unused section markers, are deleted from the generated document before saving. Paragraphs and other content are preserved, and removal is verified against the expected document text. The source template remains unchanged. Publishing still checks that an existing output file has not changed and replaces it atomically. No partial document update or cross-export result cache is introduced.

History event preparation no longer renders section bodies. The exporter renders each text-only event once, or each grouped section once after collecting its members. Text-only parts are assembled with Join. The batch shares one initialized XML template catalog with the preparation service; the service owns only catalogs it creates itself. Document replacements and output filename are prepared from the first included event and reused only after comparing the configured common fields for all later included events.

export.Word.documentContext.fields declares inputs shared by document-level expressions. Both profiles currently declare AcceptedOrderNumber, AcceptedOrderDate and Signatory, covering all six configured replacement expressions. If a document-level expression gains another input, add that input to this contract so differing values are rejected rather than silently reused.

## Timings

diagnostic.log contains Word.Export.* PERF entries. ReadEvents measures history loading; PrepareJournal measures event journal preparation. Prepare.Initialize, Prepare.Events and Prepare.RenderSections cover in-memory preparation. AcquireWord measures process creation or reuse separately from CreateDocument. ReplaceDocumentFields is separate from per-section IndexAnchors, Insert, Verify and HideMarkers. SaveDocument, CloseDocument and PublishDocument cover file work. RecordJournal and RefreshHistory cover post-generation history updates. GenerateTotal includes generation cleanup; TotalIncludingJournalAndRefresh covers the command through the final status refresh, before its final diagnostic flush.

WORD_EXPORT_BATCH records event and populated-section counts. Per-section totals include the event type in the stage name. Repeated stage names are separate samples and must be summed within one command. Measurements use the existing diagnostic timer; no estimated timings are presented as actual results.

## Manual benchmark

Use disposable copies of the output and history books. Compare 1, 10, 50 and 100 active events, first in one section and then distributed across sections. Record the first export after page initialization and an immediate repeated export separately. Retain event counts, section counts, total output characters/pages and the stage timings. A first Word launch and reuse are distinct cases; do not restart or close a user's unrelated Word process for measurements.

The existing anchor scan and full-document verification remain unchanged until timings establish their cost. Runtime Word formatting behavior and throughput require interactive verification; static contracts cannot prove them.