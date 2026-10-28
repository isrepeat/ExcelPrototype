using System;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Diagnostics;

internal static class Agregator
{
    private sealed class WorkbookSource
    {
        public string Path;
        public List<string> Sheets = new List<string>();
    }

    private sealed class Configuration
    {
        public string Path;
        public string Output;
        public string Mode;
        public List<WorkbookSource> Sources = new List<WorkbookSource>();
    }

    private static string Required(XElement element, string name)
    {
        var value = (string)element.Attribute(name);
        if (String.IsNullOrWhiteSpace(value))
            throw new InvalidDataException("Missing configuration attribute: " + element.Name + "/@" + name);
        return value;
    }

    private static string Resolve(string directory, string path)
    {
        return System.IO.Path.GetFullPath(System.IO.Path.Combine(directory, path));
    }

    private static Configuration ReadConfig(string outputOverride)
    {
        var directory = AppDomain.CurrentDomain.BaseDirectory;
        var config = new Configuration();
        config.Path = System.IO.Path.Combine(directory, "Agregator.xml");
        var document = XDocument.Load(config.Path);
        var root = document.Root;
        if (root == null || root.Name != "aggregator")
            throw new InvalidDataException("Expected aggregator configuration root.");
        config.Output = Resolve(directory, outputOverride ?? Required(root, "output"));
        config.Mode = Required(root, "copyMode");
        if (config.Mode != "values" && config.Mode != "sheet" && config.Mode != "xmlStyles")
            throw new InvalidDataException("Invalid copyMode: " + config.Mode + ". Expected values, sheet or xmlStyles.");
        var sources = root.Element("sources");
        if (sources == null || !sources.Elements("source").Any())
            throw new InvalidDataException("No source directories configured.");
        foreach (var source in sources.Elements("source"))
        {
            var sourceDirectory = Resolve(directory, Required(source, "dir"));
            if (!source.Elements("workbook").Any())
                throw new InvalidDataException("No source workbooks configured: " + sourceDirectory);
            foreach (var workbook in source.Elements("workbook"))
            {
                var item = new WorkbookSource();
                item.Path = Resolve(sourceDirectory, Required(workbook, "file"));
                item.Sheets.AddRange(workbook.Elements("sheet").Select(e => Required(e, "name")));
                if (item.Sheets.Count == 0) throw new InvalidDataException("No worksheets selected: " + item.Path);
                if (!File.Exists(item.Path)) throw new FileNotFoundException("Input workbook does not exist: " + item.Path);
                if (String.Equals(item.Path, config.Output, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Input and output paths must differ: " + item.Path);
                if (config.Mode == "xmlStyles" && !String.Equals(System.IO.Path.GetExtension(item.Path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("xmlStyles requires .xlsx sources: " + item.Path);
                config.Sources.Add(item);
            }
        }
        if (!String.Equals(System.IO.Path.GetExtension(config.Output), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Output must have the .xlsx extension: " + config.Output);
        if (File.Exists(config.Output) || Directory.Exists(config.Output))
            throw new IOException("Output already exists: " + config.Output);
        if (!Directory.Exists(System.IO.Path.GetDirectoryName(config.Output)))
            throw new DirectoryNotFoundException("Output directory does not exist: " + config.Output);
        return config;
    }

    private static object SnapshotValues(object values)
    {
        // Excel numbers arrive as Double; Int32 values represent VT_ERROR codes.
        // Rewrap them so managed COM does not write error codes as ordinary numbers.
        var cells = values as object[,];
        if (cells == null) return values is int ? new ErrorWrapper((int)values) : values;
        for (var row = cells.GetLowerBound(0); row <= cells.GetUpperBound(0); row++)
            for (var column = cells.GetLowerBound(1); column <= cells.GetUpperBound(1); column++)
                if (cells[row, column] is int) cells[row, column] = new ErrorWrapper((int)cells[row, column]);
        return cells;
    }

    private static void CopyTables(dynamic source, dynamic destination)
    {
        for (var i = 1; i <= (int)source.ListObjects.Count; i++)
        {
            dynamic original = source.ListObjects.Item(i);
            dynamic range = original.Range;
            var rows = (int)range.Rows.Count - ((bool)original.ShowTotals ? 1 : 0);
            dynamic tableRange = destination.Range[range.Address].Resize[rows, range.Columns.Count];
            dynamic table = destination.ListObjects.Add(1, tableRange, Type.Missing, 1);
            table.TableStyle = "";
            table.ShowTotals = original.ShowTotals;
            table.ShowAutoFilter = original.ShowAutoFilter;
            if ((bool)original.ShowTotals)
            {
                table.TotalsRowRange.NumberFormat = "@";
                table.TotalsRowRange.Value2 = SnapshotValues((object)original.TotalsRowRange.Value2);
                table.TotalsRowRange.ClearFormats();
            }
        }
    }

    private static void CopySheet(dynamic excel, dynamic source, dynamic result, dynamic placeholder, string sheetName, string mode, int number, string sourcePath)
    {
        var watch = Stopwatch.StartNew();
        Console.WriteLine("Copying worksheet: " + sheetName);
        dynamic original;
        try { original = source.Worksheets.Item(sheetName); }
        catch (COMException error)
        {
            throw new InvalidOperationException("Selected worksheet not found: " + sourcePath + " / " + sheetName, error);
        }
        dynamic target;
        if (mode == "sheet")
        {
            original.Copy(placeholder);
            target = result.Worksheets.Item((int)placeholder.Index - 1);
            dynamic used = target.UsedRange;
            used.Copy();
            used.PasteSpecial(12); // xlPasteValuesAndNumberFormats.
            used.FormatConditions.Delete();
            used.Validation.Delete();
            excel.CutCopyMode = false;
        }
        else
        {
            placeholder.Activate();
            target = result.Worksheets.Add();
            dynamic used = original.UsedRange;
            dynamic destination = target.Range[used.Address];
            // Protect literal strings beginning with '=' while transferring typed values.
            var stage = "setting text format";
            try
            {
                destination.NumberFormat = "@";
                stage = "assigning values";
                destination.Value2 = SnapshotValues((object)used.Value2);
                stage = "clearing the temporary text format";
                // Avoid locale-sensitive General tokens in managed COM dispatch.
                destination.ClearFormats();
            }
            catch (COMException error)
            {
                throw new InvalidOperationException("Value transfer failed while " + stage + ": " + error.Message, error);
            }
            CopyTables(original, target);
        }
        target.Visible = -1; // xlSheetVisible.
        var name = number + "_" + System.IO.Path.GetFileNameWithoutExtension(sourcePath) + "_" + sheetName;
        foreach (var character in "\\/?*[]:'") name = name.Replace(character, '_');
        target.Name = name.Substring(0, Math.Min(31, name.Length));
        Console.WriteLine("Copied in {0:F3} s", watch.Elapsed.TotalSeconds);
    }

    private static void Release(object value)
    {
        if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }

    private static void Aggregate(Configuration config)
    {
        dynamic excel = null;
        dynamic workbooks = null;
        dynamic source = null;
        dynamic result = null;
        var errors = new List<Exception>();
        // Build beside the destination, publishing only after COM and XML processing succeed.
        var temporary = config.Output + "." + Guid.NewGuid().ToString("N") + ".tmp.xlsx";
        var count = 0;
        try
        {
            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) throw new InvalidOperationException("Desktop Microsoft Excel is not installed.");
                excel = Activator.CreateInstance(excelType);
                excel.Visible = false;
                excel.ScreenUpdating = false;
                excel.DisplayAlerts = false;
                excel.EnableEvents = false;
                excel.AutomationSecurity = 3;
                workbooks = excel.Workbooks;
                result = workbooks.Add(-4167); // xlWBATWorksheet.
                excel.Calculation = -4135; // xlCalculationManual.
                excel.CalculateBeforeSave = false;
                dynamic placeholder = result.Worksheets.Item(1);
                foreach (var item in config.Sources)
                {
                    var watch = Stopwatch.StartNew();
                    Console.WriteLine("Opening: " + item.Path);
                    source = workbooks.Open(item.Path, 0, true);
                    Console.WriteLine("Opened in {0:F3} s", watch.Elapsed.TotalSeconds);
                    foreach (var sheet in item.Sheets)
                        CopySheet(excel, source, result, placeholder, sheet, config.Mode, ++count, item.Path);
                    source.Close(false);
                    Release((object)source);
                    source = null;
                }
                placeholder.Delete();
                Release((object)placeholder);
                Console.WriteLine("Saving output...");
                result.SaveAs(temporary, 51); // xlOpenXMLWorkbook.
            }
            catch (Exception error) { errors.Add(error); }
            finally
            {
                // Attempt every close operation even if an earlier cleanup step fails.
                try { if (source != null) source.Close(false); } catch (Exception error) { errors.Add(error); }
                try { if (result != null) result.Close(false); } catch (Exception error) { errors.Add(error); }
                try { if (excel != null) excel.Quit(); } catch (Exception error) { errors.Add(error); }
                Release((object)source);
                Release((object)result);
                Release((object)workbooks);
                Release((object)excel);
                // Collect short-lived RCWs from dynamic property chains after closing Excel.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (errors.Count > 0) throw new AggregateException(errors);
            if (config.Mode == "xmlStyles")
            {
                var watch = Stopwatch.StartNew();
                Console.WriteLine("Applying ordinary formatting from source XML...");
                SnapshotStyles.Apply(config.Path, temporary);
                Console.WriteLine("XML formatting applied in {0:F3} s", watch.Elapsed.TotalSeconds);
            }
            File.Move(temporary, config.Output); // Refuse to overwrite a concurrently created output.
            Console.WriteLine("Created: {0} (worksheets: {1})", config.Output, count);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        // Use UTF-8 for console output and redirected logs, including Ukrainian names.
        Console.OutputEncoding = new UTF8Encoding(false);
        var watch = Stopwatch.StartNew();
        try
        {
            var validate = false;
            string output = null;
            foreach (var argument in args)
            {
                if (argument == "--help")
                {
                    Console.WriteLine("Agregator.cmd [--validate] [output.xlsx]");
                    Console.WriteLine("Paths are relative to Agregator.xml beside the executable. Output is never overwritten.");
                    return 0;
                }
                if (argument == "--validate") validate = true;
                else if (argument.StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException("Unknown option: " + argument);
                else if (output != null) throw new ArgumentException("Specify only one optional output path.");
                else output = argument;
            }
            var config = ReadConfig(output);
            Console.WriteLine("Copy mode: " + config.Mode);
            if (validate)
            {
                Console.WriteLine("Paths validated. Input workbooks: " + config.Sources.Count);
                Console.WriteLine("Output: " + config.Output);
            }
            else Aggregate(config);
            return 0;
        }
        catch (Exception error)
        {
            var aggregate = error as AggregateException;
            if (aggregate != null)
                foreach (var cause in aggregate.Flatten().InnerExceptions) Console.Error.WriteLine("Error: " + cause.Message);
            else Console.Error.WriteLine("Error: " + error.Message);
            return 1;
        }
        finally { Console.WriteLine("Total elapsed: {0:F3} s", watch.Elapsed.TotalSeconds); }
    }
}