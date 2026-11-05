using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Diagnostics;

using ExcelTools.Interop;
using ExcelTools.Tooling;

namespace ExcelTools.Agregator {
    internal static class AggregationService {
        private static void CopyTables(dynamic source, dynamic destination) {
            for (var i = 1; i <= (int)source.ListObjects.Count; i++) {
                dynamic original = source.ListObjects.Item(i);
                dynamic range = original.Range;
                var rows = (int)range.Rows.Count - ((bool)original.ShowTotals ? 1 : 0);
                dynamic tableRange = destination.Range[range.Address].Resize[rows, range.Columns.Count];
                dynamic table = destination.ListObjects.Add(1, tableRange, Type.Missing, 1);
                table.TableStyle = "";
                table.ShowTotals = original.ShowTotals;
                table.ShowAutoFilter = original.ShowAutoFilter;
                if ((bool)original.ShowTotals) {
                    table.TotalsRowRange.NumberFormat = "@";
                    table.TotalsRowRange.Value2 = ExcelValues.PrepareForWrite((object)original.TotalsRowRange.Value2);
                    table.TotalsRowRange.ClearFormats();
                }
            }
        }

        private static string AllocateSheetName(string sheetName, string sourcePath, HashSet<string> names) {
            if (names.Add(sheetName)) {
                return sheetName;
            }
            var prefix = System.IO.Path.GetFileNameWithoutExtension(sourcePath);
            foreach (var character in "\\/?*[]:'") {
                prefix = prefix.Replace(character, '_');
            }
            var basis = prefix + "_" + sheetName;
            var candidate = basis.Substring(0, Math.Min(31, basis.Length));
            for (var suffix = 2; !names.Add(candidate); suffix++) {
                var ending = "_" + suffix;
                candidate = basis.Substring(0, Math.Min(31 - ending.Length, basis.Length)) + ending;
            }
            return candidate;
        }

        private static void CopySheet(
            dynamic excel,
            dynamic source,
            dynamic result,
            dynamic placeholder,
            string sheetName,
            string mode,
            string targetName,
            string sourcePath) {
            var watch = Stopwatch.StartNew();
            Console.WriteLine("Copying worksheet: " + sheetName);
            dynamic original;
            try {
                original = source.Worksheets.Item(sheetName);
            }
            catch (COMException error) {
                throw new InvalidOperationException("Selected worksheet not found: " + sourcePath + " / " + sheetName, error);
            }
            dynamic target;
            if (mode == "sheet") {
                original.Copy(placeholder);
                target = result.Worksheets.Item((int)placeholder.Index - 1);
                dynamic used = target.UsedRange;
                used.Copy();
                used.PasteSpecial(12); // xlPasteValuesAndNumberFormats.
                used.FormatConditions.Delete();
                used.Validation.Delete();
                excel.CutCopyMode = false;
            }
            else {
                placeholder.Activate();
                target = result.Worksheets.Add();
                dynamic used = original.UsedRange;
                dynamic destination = target.Range[used.Address];
                // Preserve strings beginning with '=' as text when transferring values.
                var stage = "setting text format";
                try {
                    destination.NumberFormat = "@";
                    stage = "assigning values";
                    destination.Value2 = ExcelValues.PrepareForWrite((object)used.Value2);
                    stage = "clearing the temporary text format";
                    // Clear formatting without relying on the localized General format name.
                    destination.ClearFormats();
                }
                catch (COMException error) {
                    throw new InvalidOperationException("Value transfer failed while " + stage + ": " + error.Message, error);
                }
                CopyTables(original, target);
            }
            target.Visible = -1; // xlSheetVisible.
            target.Name = targetName;
            Console.WriteLine("Copied in {0:F3} s", watch.Elapsed.TotalSeconds);
        }

        public static void Execute(Configuration config) {
            ExcelSession excelSession = null;
            dynamic excel = null;
            dynamic workbooks = null;
            dynamic source = null;
            dynamic result = null;
            var errors = new List<Exception>();
            // Prepare a temporary file beside the output.
            var temporary = OutputPublisher.CreateTemporaryPath(config.Output, ".xlsx");
            var count = 0;
            // Plan worksheet names before creating the temporary placeholder sheet.
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var plannedNames = new List<string>();
            foreach (var item in config.Sources) {
                foreach (var sheet in item.Sheets) {
                    plannedNames.Add(AllocateSheetName(sheet, item.Path, names));
                }
            }
            var placeholderName = "_snapshot";
            for (var suffix = 2; names.Contains(placeholderName); suffix++) {
                placeholderName = "_snapshot_" + suffix;
            }
            try {
                try {
                    excelSession = new ExcelSession();
                    excel = excelSession.Application;
                    workbooks = excelSession.Workbooks;
                    result = workbooks.Add(-4167); // xlWBATWorksheet.
                    excel.Calculation = -4135; // xlCalculationManual.
                    excel.CalculateBeforeSave = false;
                    dynamic placeholder = result.Worksheets.Item(1);
                    placeholder.Name = placeholderName;
                    foreach (var item in config.Sources) {
                        var watch = Stopwatch.StartNew();
                        Console.WriteLine("Opening: " + item.Path);
                        source = workbooks.Open(item.Path, 0, true);
                        Console.WriteLine("Opened in {0:F3} s", watch.Elapsed.TotalSeconds);
                        foreach (var sheet in item.Sheets) {
                            CopySheet(excel, source, result, placeholder, sheet, config.Mode, plannedNames[count], item.Path);
                            count++;
                        }
                        source.Close(false);
                        ComObjects.Release((object)source);
                        source = null;
                    }
                    placeholder.Delete();
                    ComObjects.Release((object)placeholder);
                    Console.WriteLine("Saving output...");
                    result.SaveAs(temporary, 51); // xlOpenXMLWorkbook.
                }
                catch (Exception error) {
                    errors.Add(error);
                }
                finally {
                    // Close all resources even if an earlier cleanup step fails.
                    try {
                        if (source != null) {
                            source.Close(false);
                        }
                    }
                    catch (Exception error) {
                        errors.Add(error);
                    }
                    try {
                        if (result != null) {
                            result.Close(false);
                        }
                    }
                    catch (Exception error) {
                        errors.Add(error);
                    }
                    try {
                        if (excelSession != null) {
                            excelSession.Dispose();
                        }
                    }
                    catch (Exception error) {
                        errors.Add(error);
                    }
                    ComObjects.Release((object)source);
                    ComObjects.Release((object)result);
                    ComObjects.CollectReleasedObjects();
                }
                if (errors.Count > 0) {
                    throw new AggregateException(errors);
                }
                if (config.Mode == "xmlStyles") {
                    var watch = Stopwatch.StartNew();
                    Console.WriteLine("Applying ordinary formatting from source XML...");
                    SnapshotStyles.Apply(config.Path, temporary);
                    Console.WriteLine("XML formatting applied in {0:F3} s", watch.Elapsed.TotalSeconds);
                }
                Console.WriteLine("Publishing output: " + config.Output);
                OutputPublisher.Publish(temporary, config.Output);
                Console.WriteLine("Created: {0} (worksheets: {1})", config.Output, count);
            }
            finally {
                if (File.Exists(temporary)) {
                    File.Delete(temporary);
                }
            }
        }

    }
}