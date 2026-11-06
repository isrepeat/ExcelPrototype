using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelTools.OpenXml {
    public static class WorkbookReader {
        internal static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        internal static XDocument ReadPart(ZipArchive zipArchive, string path) {
            var entry = zipArchive.GetEntry(path);
            if (entry == null) {
                throw new InvalidDataException("Missing workbook part: " + path);
            }
            using (var stream = entry.Open()) {
                using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) {
                    return XDocument.Load(reader);
                }
            }
        }

        private static string Text(XElement element) {
            if (element == null) {
                return "";
            }
            return String.Concat(element.Descendants(S + "t").Where(t => !t.Ancestors(S + "rPh").Any()).Select(t => t.Value));
        }

        public static ExcelTools.Tables.WorkbookSnapshot Read(string path) {
            return Read(path, null);
        }

        public static ExcelTools.Tables.WorkbookSnapshot Read(string path, IEnumerable<string> sheetNames) {
            var selectedSheets = sheetNames == null ? null : new HashSet<string>(sheetNames, StringComparer.Ordinal);
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                using (var zipArchive = new ZipArchive(stream, ZipArchiveMode.Read)) {
                    var workbook = ReadPart(zipArchive, "xl/workbook.xml");
                    if (workbook.Root.Name.Namespace != S) {
                        throw new InvalidDataException("Unsupported workbook XML namespace: " + workbook.Root.Name.NamespaceName);
                    }
                    var relations = ReadPart(zipArchive, "xl/_rels/workbook.xml.rels").Root.Elements()
                        .ToDictionary(e => (string)e.Attribute("Id"));
                    var shared = zipArchive.GetEntry("xl/sharedStrings.xml") == null
                        ? new List<string>()
                        : ReadPart(zipArchive, "xl/sharedStrings.xml").Root.Elements(S + "si").Select(Text).ToList();
                    var formats = ReadFormats(zipArchive);
                    var properties = workbook.Root.Element(S + "workbookPr");
                    var date1904 = properties != null && new[] { "1", "true" }.Contains((string)properties.Attribute("date1904"));
                    var result = new ExcelTools.Tables.WorkbookSnapshot();
                    foreach (var sheet in workbook.Root.Element(S + "sheets").Elements(S + "sheet")) {
                        if (selectedSheets != null && !selectedSheets.Contains((string)sheet.Attribute("name"))) {
                            continue;
                        }
                        XElement relation;
                        if (!relations.TryGetValue((string)sheet.Attribute(R + "id"), out relation)
                            || (string)relation.Attribute("TargetMode") == "External") {
                            throw new InvalidDataException("Invalid worksheet relationship: " + (string)sheet.Attribute("name"));
                        }
                        if (!((string)relation.Attribute("Type") ?? "").EndsWith("/worksheet", StringComparison.Ordinal)) {
                            throw new InvalidDataException("Only worksheets are supported: " + (string)sheet.Attribute("name"));
                        }
                        var part = ResolvePart(relation);
                        result.Sheets.Add(ReadSheet(zipArchive, part, (string)sheet.Attribute("name"), shared, formats, date1904));
                    }
                    if (selectedSheets != null) {
                        foreach (var name in selectedSheets) {
                            if (result.FindSheet(name) == null) {
                                throw new InvalidDataException("Selected worksheet '" + name + "' does not exist in workbook: " + path);
                            }
                        }
                    }
                    return result;
                }
            }
        }

        internal static string ResolvePart(XElement relation) {
            var uri = new Uri(new Uri("http://workbook/xl/workbook.xml"), (string)relation.Attribute("Target"));
            return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        }

        private static List<string> ReadFormats(ZipArchive zipArchive) {
            if (zipArchive.GetEntry("xl/styles.xml") == null) {
                return new List<string> { "" };
            }
            var root = ReadPart(zipArchive, "xl/styles.xml").Root;
            var custom = root.Element(S + "numFmts");
            var definitions = custom == null ? new Dictionary<int, string>()
                : custom.Elements(S + "numFmt").ToDictionary(e => (int)e.Attribute("numFmtId"), e => (string)e.Attribute("formatCode"));
            var result = new List<string>();
            var styles = root.Element(S + "cellXfs");
            if (styles == null) {
                return new List<string> { "" };
            }
            foreach (var style in styles.Elements(S + "xf")) {
                var id = (int?)style.Attribute("numFmtId") ?? 0;
                string format;
                if (!definitions.TryGetValue(id, out format)) {
                    format = (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 50 && id <= 58)
                        ? "yyyy-mm-dd hh:mm:ss" : id >= 45 && id <= 47 ? "hh:mm:ss" : "";
                    if (id == 46) {
                        format = "[h]:mm:ss";
                    }
                }
                result.Add(format);
            }
            return result;
        }

        private static ExcelTools.Tables.SheetSnapshot ReadSheet(
            ZipArchive zipArchive,
            string part,
            string name,
            List<string> shared,
            List<string> formats,
            bool date1904) {
            var result = new ExcelTools.Tables.SheetSnapshot { Name = name };
            var rows = ReadPart(zipArchive, part).Root.Element(S + "sheetData");
            if (rows == null) {
                return result;
            }
            var sharedFormulas = new Dictionary<string, XElement>();
            foreach (var cell in rows.Descendants(S + "c")) {
                var formula = cell.Element(S + "f");
                if (formula != null && (string)formula.Attribute("t") == "shared" && formula.Value.Length > 0) {
                    sharedFormulas[(string)formula.Attribute("si")] = cell;
                }
            }
            var previousRow = 0;
            // Read every saved row, regardless of worksheet/table filters and hidden state.
            foreach (var row in rows.Elements(S + "row")) {
                var number = (int?)row.Attribute("r") ?? previousRow + 1;
                previousRow = number;
                var values = new List<string>();
                foreach (var cell in row.Elements(S + "c")) {
                    if (cell.Element(S + "v") == null && cell.Element(S + "f") == null && cell.Element(S + "is") == null) {
                        continue;
                    }
                    var address = (string)cell.Attribute("r");
                    if (String.IsNullOrEmpty(address)) {
                        throw new InvalidDataException("Missing cell address in sheet: " + name);
                    }
                    var column = ExcelTools.Tables.CellAddress.ColumnNumber(address);
                    result.Width = Math.Max(result.Width, column);
                    while (values.Count < column) {
                        values.Add("");
                    }
                    values[column - 1] = ReadValue(cell, shared, formats, date1904, sharedFormulas);
                }
                if (values.Any(v => v.Length > 0)) {
                    result.Rows.Add(new ExcelTools.Tables.RowSnapshot { Number = number, Values = values.ToArray() });
                }
            }
            return result;
        }

        private static string ReadValue(
            XElement cell,
            List<string> shared,
            List<string> formats,
            bool date1904,
            Dictionary<string, XElement> sharedFormulas) {
            var formula = cell.Element(S + "f");
            if (formula != null) {
                if ((string)formula.Attribute("t") == "dataTable") {
                    throw new InvalidDataException("Data-table formulas are not supported: " + (string)cell.Attribute("r"));
                }
                var expression = formula.Value;
                if (expression.Length == 0 && (string)formula.Attribute("t") == "shared") {
                    XElement original;
                    if (!sharedFormulas.TryGetValue((string)formula.Attribute("si"), out original)) {
                        throw new InvalidDataException("Missing shared formula: " + (string)cell.Attribute("r"));
                    }
                    expression = FormulaTranslator.Translate(
                        original.Element(S + "f").Value,
                        (string)original.Attribute("r"),
                        (string)cell.Attribute("r")
                    );
                }
                return "=" + expression;
            }
            var type = (string)cell.Attribute("t") ?? "n";
            var text = (string)cell.Element(S + "v") ?? "";
            if (type == "inlineStr") {
                return Text(cell.Element(S + "is"));
            }
            if (type == "s") {
                return shared[Int32.Parse(text, CultureInfo.InvariantCulture)];
            }
            if (type == "b") {
                return text == "1" ? "True" : "False";
            }
            if (type == "d") {
                return FormatDate(DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind));
            }
            if (type != "n" || text.Length == 0) {
                return text;
            }
            var number = Double.Parse(text, CultureInfo.InvariantCulture);
            var style = (int?)cell.Attribute("s") ?? 0;
            if (style >= formats.Count) {
                throw new InvalidDataException("Invalid cell style: " + style);
            }
            var format = formats[style].Split(';')[0];
            var duration = Regex.IsMatch(format, @"\[(h+|m+|s+)\]", RegexOptions.IgnoreCase);
            format = Regex.Replace(format, @"""[^""]*""|\\.|\[[^\]]*\]", "");
            if (duration) {
                var time = TimeSpan.FromMilliseconds(Math.Round(number * 86400000));
                return (time.Days == 0 ? "" : time.Days + (time.Days == 1 ? " day, " : " days, "))
                    + (time.Days == 0 ? ((long)time.TotalHours).ToString(CultureInfo.InvariantCulture) : time.Hours.ToString(CultureInfo.InvariantCulture))
                    + ":" + time.Minutes.ToString("00") + ":" + time.Seconds.ToString("00")
                    + (time.Milliseconds == 0 ? "" : "." + (time.Milliseconds * 1000).ToString("000000"));
            }
            if (Regex.IsMatch(format, @"[dmhys]", RegexOptions.IgnoreCase)) {
                if (number >= 0 && number < 1) {
                    return DateTime.Today.AddMilliseconds(Math.Round(number * 86400000)).ToString("HH:mm:ss", CultureInfo.InvariantCulture)
                        + Fraction(number);
                }
                var adjusted = number;
                if (!date1904 && number > 0 && number < 60) {
                    adjusted++;
                }
                var date = date1904 ? new DateTime(1904, 1, 1).AddMilliseconds(Math.Round(number * 86400000))
                    : DateTime.FromOADate(adjusted);
                return FormatDate(date);
            }
            return FormatNumber(text, number);
        }

        private static string FormatNumber(string text, double number) {
            // Preserve integer precision and Python's lowercase scientific notation.
            if (Regex.IsMatch(text, @"^[+-]?[0-9]+$")) {
                var digits = text.TrimStart('+', '-').TrimStart('0');
                if (digits.Length == 0) {
                    return "0";
                }
                return (text.StartsWith("-", StringComparison.Ordinal) ? "-" : "") + digits;
            }
            var result = number.ToString("R", CultureInfo.InvariantCulture).ToLowerInvariant();
            if (result.IndexOf('e') < 0 && number == Math.Truncate(number)) {
                return result + ".0";
            }
            return result;
        }

        private static string Fraction(double serial) {
            var milliseconds = ((long)Math.Round(serial * 86400000)) % 1000;
            return milliseconds == 0 ? "" : "." + (milliseconds * 1000).ToString("000000");
        }

        private static string FormatDate(DateTime dateTime) {
            return dateTime.TimeOfDay == TimeSpan.Zero ? dateTime.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture)
                : dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                    + (dateTime.Millisecond == 0 ? "" : "." + (dateTime.Millisecond * 1000).ToString("000000"));
        }
    }
}