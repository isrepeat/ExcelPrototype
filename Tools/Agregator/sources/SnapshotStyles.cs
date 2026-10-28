using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Globalization;

// Merge ordinary Open XML formatting into COM-generated value snapshots in-process.
internal static class SnapshotStyles
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static XDocument Read(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path);
        if (entry == null) throw new InvalidDataException("Missing package part: " + path);
        using (var stream = entry.Open()) return XDocument.Load(stream);
    }

    private static void Write(ZipArchive zip, string path, XDocument document)
    {
        var old = zip.GetEntry(path);
        if (old != null) old.Delete();
        var entry = zip.CreateEntry(path, CompressionLevel.Fastest);
        using (var stream = entry.Open()) document.Save(stream, SaveOptions.DisableFormatting);
    }

    private static string Related(ZipArchive zip, string part, string id)
    {
        var slash = part.LastIndexOf('/');
        var directory = part.Substring(0, slash + 1);
        var relationships = Read(zip, directory + "_rels/" + part.Substring(slash + 1) + ".rels");
        var relationship = relationships.Root.Elements().Single(e => (string)e.Attribute("Id") == id);
        if ((string)relationship.Attribute("TargetMode") == "External")
            throw new InvalidDataException("External package relationship is not supported: " + part);
        return new Uri(new Uri("http://package/" + part), (string)relationship.Attribute("Target")).AbsolutePath.TrimStart('/');
    }

    private static string Sheet(ZipArchive zip, string name)
    {
        var workbook = Read(zip, "xl/workbook.xml");
        var sheet = workbook.Root.Element(S + "sheets").Elements().Single(e => (string)e.Attribute("name") == name);
        return Related(zip, "xl/workbook.xml", (string)sheet.Attribute(R + "id"));
    }

    // Resolve source theme colors before combining styles from different workbooks.
    private static void ResolveTheme(ZipArchive zip, XElement styles)
    {
        var workbookRels = Read(zip, "xl/_rels/workbook.xml.rels");
        var themeRel = workbookRels.Root.Elements().FirstOrDefault(e => ((string)e.Attribute("Type")).EndsWith("/theme"));
        if (themeRel == null) return;
        var theme = Read(zip, Related(zip, "xl/workbook.xml", (string)themeRel.Attribute("Id")));
        var scheme = theme.Descendants(A + "clrScheme").Single();
        var order = new[] { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
        foreach (var color in styles.Descendants().Where(e => e.Attribute("theme") != null))
        {
            var index = (int)color.Attribute("theme");
            var definition = scheme.Element(A + order[index]).Elements().First();
            var rgb = (string)definition.Attribute("lastClr") ?? (string)definition.Attribute("val");
            color.Attribute("theme").Remove();
            color.SetAttributeValue("rgb", "FF" + rgb);
        }
        var indexed = styles.Element(S + "colors");
        if (indexed != null && indexed.Element(S + "indexedColors") != null)
        {
            var palette = indexed.Element(S + "indexedColors").Elements().ToList();
            foreach (var color in styles.Descendants().Where(e => e.Attribute("indexed") != null))
            {
                var index = (int)color.Attribute("indexed");
                if (index < palette.Count)
                {
                    color.SetAttributeValue("rgb", (string)palette[index].Attribute("rgb"));
                    color.Attribute("indexed").Remove();
                }
            }
        }
        var fontScheme = theme.Descendants(A + "fontScheme").Single();
        foreach (var font in styles.Descendants(S + "font"))
        {
            var schemeNode = font.Element(S + "scheme");
            if (schemeNode == null) continue;
            var kind = (string)schemeNode.Attribute("val");
            if (kind != "major" && kind != "minor") continue;
            var family = (string)fontScheme.Element(A + (kind + "Font")).Element(A + "latin").Attribute("typeface");
            if (!String.IsNullOrEmpty(family))
            {
                if (font.Element(S + "name") == null) font.AddFirst(new XElement(S + "name"));
                font.Element(S + "name").SetAttributeValue("val", family);
            }
        }
        // Explicit font names remain valid without following the output workbook theme.
        foreach (var schemeNode in styles.Descendants(S + "scheme").ToList()) schemeNode.Remove();
    }

    private sealed class StyleMap
    {
        public int CellOffset;
        public int DxfOffset;
        public double ColumnScale;
        public Dictionary<string, string> TableNames = new Dictionary<string, string>();
    }

    private static XElement Collection(XElement root, string name)
    {
        var result = root.Element(S + name);
        if (result == null) { result = new XElement(S + name); root.Add(result); }
        return result;
    }

    private static int Append(XElement destination, XElement source)
    {
        var offset = destination.Elements().Count();
        if (source != null) destination.Add(source.Elements().Select(e => new XElement(e)));
        destination.SetAttributeValue("count", destination.Elements().Count());
        return offset;
    }

    private static void Offset(XElement element, string attribute, int offset)
    {
        var value = element.Attribute(attribute);
        if (value != null) value.Value = ((int)value + offset).ToString();
    }

    private static int DigitWidth(XElement font)
    {
        using (var measured = new Font((string)font.Element(S + "name").Attribute("val"),
            (float)font.Element(S + "sz").Attribute("val"), FontStyle.Regular, GraphicsUnit.Point))
        {
            var width = 0;
            foreach (var digit in "0123456789")
                width = Math.Max(width, TextRenderer.MeasureText(digit.ToString(), measured, Size.Empty, TextFormatFlags.NoPadding).Width);
            return width;
        }
    }

    private static StyleMap MergeStyles(ZipArchive input, XElement output, int sourceNumber, ref int outputDigitWidth)
    {
        var original = Read(input, "xl/styles.xml").Root;
        ResolveTheme(input, original);
        var normalStyle = original.Element(S + "cellStyles").Elements().First(e => (int?)e.Attribute("builtinId") == 0);
        var normalXf = original.Element(S + "cellStyleXfs").Elements().ElementAt((int)normalStyle.Attribute("xfId"));
        var normalFont = original.Element(S + "fonts").Elements().ElementAt((int?)normalXf.Attribute("fontId") ?? 0);
        var sourceDigitWidth = DigitWidth(normalFont);
        if (outputDigitWidth == 0)
        {
            // Column widths depend on the workbook's Normal font, not the cell font.
            output.Element(S + "fonts").Elements().First().ReplaceWith(new XElement(normalFont));
            outputDigitWidth = sourceDigitWidth;
        }
        var numFormats = Collection(output, "numFmts");
        var nextNumber = Math.Max(163, numFormats.Elements().Select(e => (int)e.Attribute("numFmtId")).DefaultIfEmpty(163).Max()) + 1;
        var numberMap = new Dictionary<int, int>();
        var sourceNumbers = original.Element(S + "numFmts");
        if (sourceNumbers != null)
            foreach (var format in sourceNumbers.Elements())
            {
                var copy = new XElement(format);
                numberMap[(int)format.Attribute("numFmtId")] = nextNumber;
                copy.SetAttributeValue("numFmtId", nextNumber++);
                numFormats.Add(copy);
            }
        numFormats.SetAttributeValue("count", numFormats.Elements().Count());
        foreach (var format in original.Descendants().Where(e => e.Attribute("numFmtId") != null))
        {
            int mapped;
            if (numberMap.TryGetValue((int)format.Attribute("numFmtId"), out mapped)) format.SetAttributeValue("numFmtId", mapped);
        }
        var fonts = Append(Collection(output, "fonts"), original.Element(S + "fonts"));
        var fills = Append(Collection(output, "fills"), original.Element(S + "fills"));
        var borders = Append(Collection(output, "borders"), original.Element(S + "borders"));
        var baseOffset = Collection(output, "cellStyleXfs").Elements().Count();
        foreach (var xf in original.Descendants(S + "xf"))
        {
            Offset(xf, "fontId", fonts);
            // Fill IDs 0 (none) and 1 (gray125) are reserved. Duplicating "none"
            // at another index can turn inherited/no-fill cells into explicit white fills.
            var fillId = xf.Attribute("fillId");
            if (fillId != null && (int)fillId > 1) Offset(xf, "fillId", fills);
            Offset(xf, "borderId", borders);
            Offset(xf, "xfId", baseOffset);
        }
        Append(Collection(output, "cellStyleXfs"), original.Element(S + "cellStyleXfs"));
        var map = new StyleMap();
        map.ColumnScale = (double)sourceDigitWidth / outputDigitWidth;
        map.CellOffset = Append(Collection(output, "cellXfs"), original.Element(S + "cellXfs"));
        map.DxfOffset = Append(Collection(output, "dxfs"), original.Element(S + "dxfs"));
        var tableStyles = Collection(output, "tableStyles");
        var sourceTables = original.Element(S + "tableStyles");
        if (sourceTables != null)
            foreach (var tableStyle in sourceTables.Elements(S + "tableStyle"))
            {
                var copy = new XElement(tableStyle);
                var name = "Snapshot" + sourceNumber + "_" + (string)copy.Attribute("name");
                map.TableNames[(string)copy.Attribute("name")] = name;
                copy.SetAttributeValue("name", name);
                foreach (var element in copy.Elements()) Offset(element, "dxfId", map.DxfOffset);
                tableStyles.Add(copy);
            }
        tableStyles.SetAttributeValue("count", tableStyles.Elements().Count());
        return map;
    }

    private static void CopyElement(XElement original, XElement target, string name)
    {
        target.Elements(S + name).Remove();
        var element = original.Element(S + name);
        if (element != null) target.Add(new XElement(element));
    }

    private static void ScaleWidth(XElement element, string attribute, double scale)
    {
        var value = element.Attribute(attribute);
        if (value == null || scale == 1) return;
        var width = Double.Parse(value.Value, CultureInfo.InvariantCulture);
        value.Value = (Math.Round(width * scale * 256) / 256).ToString(CultureInfo.InvariantCulture);
    }

    private static void ApplySheet(ZipArchive input, string sourcePath, ZipArchive output, string outputPath, StyleMap map)
    {
        var original = Read(input, sourcePath).Root;
        var document = Read(output, outputPath);
        var target = document.Root;
        foreach (var name in new[] { "sheetFormatPr", "cols", "mergeCells", "sheetViews", "pageMargins", "printOptions", "headerFooter" })
            CopyElement(original, target, name);
        var columns = target.Element(S + "cols");
        if (columns != null)
            foreach (var col in columns.Elements())
            {
                Offset(col, "style", map.CellOffset);
                ScaleWidth(col, "width", map.ColumnScale);
            }
        var sheetFormat = target.Element(S + "sheetFormatPr");
        if (sheetFormat != null) ScaleWidth(sheetFormat, "defaultColWidth", map.ColumnScale);
        var defaultRowStyle = new Dictionary<int, int>();
        var sourceColumns = original.Element(S + "cols");
        if (sourceColumns != null)
            foreach (var col in sourceColumns.Elements())
                if (col.Attribute("style") != null)
                    for (var c = (int)col.Attribute("min"); c <= (int)col.Attribute("max"); c++)
                        defaultRowStyle[c] = (int)col.Attribute("style");
        var data = target.Element(S + "sheetData");
        var targetRows = data.Elements(S + "row").ToDictionary(e => (string)e.Attribute("r"));
        var sourceRows = original.Element(S + "sheetData").Elements(S + "row").ToDictionary(e => (string)e.Attribute("r"));
        // COM may materialize blank cells that originally inherited column/row styles.
        foreach (var targetRow in targetRows.Values)
        {
            XElement sourceRow;
            sourceRows.TryGetValue((string)targetRow.Attribute("r"), out sourceRow);
            foreach (var cell in targetRow.Elements(S + "c"))
            {
                int inherited;
                if (!defaultRowStyle.TryGetValue(Column((string)cell.Attribute("r")), out inherited)) inherited = 0;
                if (sourceRow != null) inherited = (int?)sourceRow.Attribute("s") ?? inherited;
                cell.SetAttributeValue("s", inherited + map.CellOffset);
            }
        }
        foreach (var row in original.Element(S + "sheetData").Elements(S + "row"))
        {
            var key = (string)row.Attribute("r");
            XElement targetRow;
            if (!targetRows.TryGetValue(key, out targetRow)) { targetRow = new XElement(S + "row"); data.Add(targetRow); }
            targetRow.ReplaceAttributes(row.Attributes());
            Offset(targetRow, "s", map.CellOffset);
            var cells = targetRow.Elements(S + "c").ToDictionary(e => (string)e.Attribute("r"));
            foreach (var cell in row.Elements(S + "c"))
            {
                var address = (string)cell.Attribute("r");
                XElement destination;
                if (!cells.TryGetValue(address, out destination))
                {
                    destination = new XElement(S + "c", new XAttribute("r", address));
                    targetRow.Add(destination);
                }
                int inherited;
                if (!defaultRowStyle.TryGetValue(Column(address), out inherited)) inherited = 0;
                inherited = (int?)row.Attribute("s") ?? inherited;
                destination.SetAttributeValue("s", ((int?)cell.Attribute("s") ?? inherited) + map.CellOffset);
            }
            // Source blank cells may be absent from the value snapshot; restore cell order.
            targetRow.ReplaceNodes(targetRow.Elements(S + "c").OrderBy(e => Column((string)e.Attribute("r"))).ToList());
        }
        data.ReplaceNodes(data.Elements(S + "row").OrderBy(e => (int)e.Attribute("r")).ToList());
        var sourceParts = original.Element(S + "tableParts");
        var targetParts = target.Element(S + "tableParts");
        if (sourceParts != null && targetParts != null)
        {
            var originals = sourceParts.Elements().Select(e => Read(input, Related(input, sourcePath, (string)e.Attribute(R + "id"))).Root).ToList();
            foreach (var part in targetParts.Elements())
            {
                var tablePath = Related(output, outputPath, (string)part.Attribute(R + "id"));
                var tableDocument = Read(output, tablePath);
                var table = tableDocument.Root;
                var sourceTable = originals.Single(e => (string)e.Attribute("ref") == (string)table.Attribute("ref"));
                CopyElement(sourceTable, table, "tableStyleInfo");
                var info = table.Element(S + "tableStyleInfo");
                string mapped;
                if (info != null && map.TableNames.TryGetValue((string)info.Attribute("name") ?? "", out mapped)) info.SetAttributeValue("name", mapped);
                foreach (var attribute in new[] { "headerRowDxfId", "dataDxfId", "totalsRowDxfId" })
                {
                    var value = sourceTable.Attribute(attribute);
                    if (value != null) table.SetAttributeValue(attribute, (int)value + map.DxfOffset);
                }
                var originalColumns = sourceTable.Element(S + "tableColumns").Elements().ToList();
                var destinationColumns = table.Element(S + "tableColumns").Elements().ToList();
                for (var i = 0; i < originalColumns.Count; i++)
                    foreach (var attribute in new[] { "headerRowDxfId", "dataDxfId", "totalsRowDxfId" })
                    {
                        var value = originalColumns[i].Attribute(attribute);
                        if (value != null) destinationColumns[i].SetAttributeValue(attribute, (int)value + map.DxfOffset);
                    }
                Write(output, tablePath, tableDocument);
            }
        }
        // Keep worksheet children in the order required by SpreadsheetML.
        var order = "sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems tableParts extLst".Split(' ');
        target.ReplaceNodes(target.Elements().OrderBy(e => Array.IndexOf(order, e.Name.LocalName)).ToList());
        Write(output, outputPath, document);
    }

    private static int Column(string address)
    {
        var result = 0;
        foreach (var ch in address) { if (ch < 'A' || ch > 'Z') break; result = result * 26 + ch - 'A' + 1; }
        return result;
    }

    public static void Apply(string configPath, string outputPath)
    {
        var config = XDocument.Load(configPath);
        var directory = Path.GetDirectoryName(Path.GetFullPath(configPath));
        var temporary = outputPath + ".styles-" + Guid.NewGuid().ToString("N") + ".tmp";
        File.Copy(outputPath, temporary);
        try
        {
            using (var output = ZipFile.Open(temporary, ZipArchiveMode.Update))
            {
                var styles = Read(output, "xl/styles.xml");
                var outputSheets = Read(output, "xl/workbook.xml").Root.Element(S + "sheets").Elements().Select(e => (string)e.Attribute("name")).ToList();
                var index = 0;
                var sourceNumber = 0;
                var outputDigitWidth = 0;
                foreach (var source in config.Root.Element("sources").Elements("source"))
                    foreach (var workbook in source.Elements("workbook"))
                    {
                        var path = Path.GetFullPath(Path.Combine(directory, (string)source.Attribute("dir"), (string)workbook.Attribute("file")));
                        using (var input = ZipFile.OpenRead(path))
                        {
                            var map = MergeStyles(input, styles.Root, ++sourceNumber, ref outputDigitWidth);
                            foreach (var sheet in workbook.Elements("sheet"))
                            {
                                var name = (string)sheet.Attribute("name");
                                Console.WriteLine("Applying XML styles: sheet " + (index + 1));
                                ApplySheet(input, Sheet(input, name), output, Sheet(output, outputSheets[index++]), map);
                            }
                        }
                    }
                var order = "numFmts fonts fills borders cellStyleXfs cellXfs cellStyles dxfs tableStyles colors extLst".Split(' ');
                styles.Root.ReplaceNodes(styles.Root.Elements().OrderBy(e => Array.IndexOf(order, e.Name.LocalName)).ToList());
                Write(output, "xl/styles.xml", styles);
            }
            File.Replace(temporary, outputPath, null);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}