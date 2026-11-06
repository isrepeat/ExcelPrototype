using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;

namespace ExcelTools.OpenXml {
    public sealed class TextCellEdit {
        public string Sheet;
        public int Row;
        public int Column;
        public string Value;
    }

    public static class WorkbookWriter {
        private static readonly XNamespace S = WorkbookReader.S;

        public static void ApplyTextEdits(string stagedWorkbook, IEnumerable<TextCellEdit> edits, string appendTable = null) {
            using (var stream = File.Open(stagedWorkbook, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Update)) {
                    const string workbookPath = "xl/workbook.xml";
                    var workbook = WorkbookReader.ReadPart(archive, workbookPath);
                    var relationships = WorkbookReader.ReadPart(archive, "xl/_rels/workbook.xml.rels").Root.Elements()
                        .ToDictionary(element => (string)element.Attribute("Id"));
                    var sheets = workbook.Root.Element(S + "sheets").Elements(S + "sheet")
                        .ToDictionary(element => (string)element.Attribute("name"), StringComparer.Ordinal);

                    foreach (var group in edits.GroupBy(edit => edit.Sheet, StringComparer.Ordinal)) {
                        var relation = relationships[(string)sheets[group.Key].Attribute(WorkbookReader.R + "id")];
                        var path = WorkbookReader.ResolvePart(relation);
                        var document = WorkbookReader.ReadPart(archive, path);
                        if (appendTable != null) {
                            ExtendTable(archive, path, document.Root, appendTable, group);
                        }
                        ApplySheetEdits(document.Root, group);
                        WritePart(archive, path, document);
                    }
                    var calculation = workbook.Root.Element(S + "calcPr");
                    if (calculation != null) {
                        calculation.SetAttributeValue("fullCalcOnLoad", "1");
                        WritePart(archive, workbookPath, workbook);
                    }
                }
            }
        }

        private static void ExtendTable(
            ZipArchive archive, string sheetPart, XElement sheet, string tableName, IEnumerable<TextCellEdit> edits) {
            var path = WorkbookReader.TablePart(archive, sheetPart, tableName);
            var table = WorkbookReader.ReadPart(archive, path);
            var range = ((string)table.Root.Attribute("ref")).Split(':');
            var last = range[1];
            var endRow = ExcelTools.Tables.CellAddress.RowNumber(last);
            var endColumn = ExcelTools.Tables.CellAddress.ColumnNumber(last);
            var startColumn = ExcelTools.Tables.CellAddress.ColumnNumber(range[0]);
            var lastEdit = edits.Max(edit => edit.Row);
            if (lastEdit <= endRow) {
                return;
            }
            if (((int?)table.Root.Attribute("totalsRowCount") ?? 0) != 0 || lastEdit > 1048576
                || edits.Any(edit => edit.Column < startColumn || edit.Column > endColumn)) {
                throw new InvalidOperationException("Cannot extend table: " + tableName);
            }
            var sheetData = sheet.Element(S + "sheetData");
            var existing = sheetData.Elements(S + "row").Where(row => (int)row.Attribute("r") > endRow)
                .ToDictionary(row => (int)row.Attribute("r"));
            for (var number = endRow + 1; number <= lastEdit; number++) {
                XElement row;
                if (existing.TryGetValue(number, out row)) {
                    if (row.Elements(S + "c").Any(cell => {
                        var column = ExcelTools.Tables.CellAddress.ColumnNumber((string)cell.Attribute("r"));
                        return column >= startColumn && column <= endColumn && cell.HasElements;
                    })) {
                        throw new InvalidOperationException("Occupied row below table: " + number);
                    }
                }
                else {
                    row = new XElement(S + "row", new XAttribute("r", number));
                    var next = sheetData.Elements(S + "row").FirstOrDefault(item => (int)item.Attribute("r") > number);
                    if (next == null) {
                        sheetData.Add(row);
                    }
                    else {
                        next.AddBeforeSelf(row);
                    }
                }
            }
            var expanded = range[0] + ":" + ExcelTools.Tables.CellAddress.ColumnName(endColumn) + lastEdit;
            table.Root.SetAttributeValue("ref", expanded);
            var filter = table.Root.Element(S + "autoFilter");
            if (filter != null) {
                filter.SetAttributeValue("ref", expanded);
            }
            var dimension = sheet.Element(S + "dimension");
            if (dimension != null) {
                var bounds = ((string)dimension.Attribute("ref")).Split(':');
                var lastCell = bounds[bounds.Length - 1];
                var column = Math.Max(endColumn, ExcelTools.Tables.CellAddress.ColumnNumber(lastCell));
                var row = Math.Max(lastEdit, ExcelTools.Tables.CellAddress.RowNumber(lastCell));
                dimension.SetAttributeValue("ref", bounds[0] + ":" + ExcelTools.Tables.CellAddress.ColumnName(column) + row);
            }
            WritePart(archive, path, table);
        }

        private static void ApplySheetEdits(XElement sheet, IEnumerable<TextCellEdit> edits) {
            if (sheet.Element(S + "sheetProtection") != null) {
                throw new InvalidOperationException("Cannot edit a protected worksheet.");
            }
            var rows = sheet.Element(S + "sheetData").Elements(S + "row")
                .ToDictionary(row => (int)row.Attribute("r"));
            var blockedRanges = sheet.Descendants(S + "mergeCell").Select(element => (string)element.Attribute("ref"))
                .Concat(sheet.Descendants(S + "f").Where(formula => (string)formula.Attribute("t") == "array"
                    || (string)formula.Attribute("t") == "dataTable").Select(formula => (string)formula.Attribute("ref")))
                .Where(range => !String.IsNullOrEmpty(range)).ToArray();
            foreach (var edit in edits) {
                var address = ExcelTools.Tables.CellAddress.ColumnName(edit.Column) + edit.Row;
                var row = rows[edit.Row];
                var cell = row.Elements(S + "c").FirstOrDefault(element => (string)element.Attribute("r") == address);
                if ((cell != null && (cell.Element(S + "f") != null || cell.Attribute("cm") != null || cell.Attribute("vm") != null))
                    || blockedRanges.Any(range => ContainsCell(range, edit.Row, edit.Column))) {
                    throw new InvalidOperationException("Cannot replace a formula, merged or special cell: " + edit.Sheet + "!" + address);
                }
                if (cell == null) {
                    cell = new XElement(S + "c", new XAttribute("r", address));
                    var next = row.Elements(S + "c").FirstOrDefault(element =>
                        ExcelTools.Tables.CellAddress.ColumnNumber((string)element.Attribute("r")) > edit.Column);
                    if (next == null) {
                        var previous = row.Elements(S + "c").LastOrDefault();
                        if (previous == null) {
                            row.AddFirst(cell);
                        }
                        else {
                            previous.AddAfterSelf(cell);
                        }
                    }
                    else {
                        next.AddBeforeSelf(cell);
                    }
                }
                cell.SetAttributeValue("t", "inlineStr");
                cell.Elements(S + "v").Remove();
                cell.Elements(S + "is").Remove();
                cell.AddFirst(new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                    edit.Value)));
            }
        }

        private static bool ContainsCell(string range, int row, int column) {
            var endpoints = range.Split(':');
            var first = endpoints[0].Replace("$", "");
            var last = endpoints[endpoints.Length - 1].Replace("$", "");
            return column >= ExcelTools.Tables.CellAddress.ColumnNumber(first) && column <= ExcelTools.Tables.CellAddress.ColumnNumber(last)
                && row >= Int32.Parse(first.TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()))
                && row <= Int32.Parse(last.TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()));
        }

        private static void WritePart(ZipArchive archive, string path, XDocument document) {
            archive.GetEntry(path).Delete();
            using (var stream = archive.CreateEntry(path).Open()) {
                document.Save(stream, SaveOptions.DisableFormatting);
            }
        }
    }
}