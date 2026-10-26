# Reporter acting assignments

Reporter acting status requires both an acceptance order (`Прийняття.Наказ`) and its date (`Прийняття`). An admission order alone does not start duties and does not decorate the reporter as acting. The reporter reader uses the same acceptance predicate as the Movement chain reader and skips unaccepted rows before reading their acceptance dates. A supplied acceptance order without a date remains invalid.

Reporter positions come from the Movement workbook's `ТВО` worksheet, configured under `reporter.tvo.source` in PROD and TEST. The source joins the shared dependency snapshot and uses its name index; changing a report date filters the loaded records without refreshing the source workbook.

An assignment applies when `Прийняття <= IncomingDate` and `Здавання` is empty or later than `IncomingDate`. The start date is inclusive and the actual handover date is exclusive. `Здавання.План` does not terminate an assignment. Names are normalized with the existing table-query rules. Multiple active assignments for the same name or an active assignment without `Посада.Заміщує` stop processing with a configured diagnostic.

Reporter candidates use the existing incoming date, or the mode's configured order-date autofill when the date is empty. The main form recalculates the reporter code when the incoming date changes with autofill enabled. Main and additional report exports resolve the assignment again using their own incoming dates. Outside an acting interval they resolve the staff code from Personnel, preventing an earlier acting code from surviving a date change. Only active assignments receive the existing `reporter.tvo.prefix` in preview and export.

Personnel-based TVO chains for event owners and arrival restoration retain their existing sources. The former undated reporter-position map from Personnel is no longer prepared or consumed.

Static validation covers configuration imports, snapshot ownership, source headers, VBA declarations and blocks, source-size limits and page markup. Excel/VBA execution requires separate authorization. Runtime acceptance should cover dates before, on and after acceptance; empty versus actual handover dates; ignored planned handover; repeated nonoverlapping assignments; overlapping assignments; and moving a report date out of an acting interval.