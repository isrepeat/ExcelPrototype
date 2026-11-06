using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ExcelTools.TableDiff {
    internal static class Program {
        private static int Main(string[] args) {
            return ExcelTools.Tooling.Console.ConsoleTool.Run(args, "TableDiff.log", delegate {
                if (args.Length == 3 && args[0] == "--prepared-report" && args[2] == "--no-pause") {
                    WritePreparedReport(args[1]);
                    return;
                }
                string configFile = null;
                string profileName = null;
                string oldFile = null;
                string newFile = null;
                var validate = false;
                for (var index = 0; index < args.Length; index++) {
                    var argument = args[index];
                    if (argument == "--help") {
                        Console.WriteLine("TableDiff.exe [--config file.xml] [--profile name | old.xlsx new.xlsx] [--validate] [--no-pause]");
                        Console.WriteLine("Config names are searched beside the executable, in resources, then in ../resources.");
                        return;
                    }
                    else if (argument == "--no-pause") {
                    }
                    else if (argument == "--validate") {
                        validate = true;
                    }
                    else if (argument == "--config") {
                        if (configFile != null || ++index >= args.Length || String.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("--", StringComparison.Ordinal)) {
                            throw new ArgumentException("Specify --config once followed by an XML file.");
                        }
                        configFile = args[index];
                    }
                    else if (argument == "--profile") {
                        if (profileName != null || ++index >= args.Length || String.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("--", StringComparison.Ordinal)) {
                            throw new ArgumentException("Specify --profile once followed by a profile name.");
                        }
                        profileName = args[index];
                    }
                    else if (argument.StartsWith("--", StringComparison.Ordinal)) {
                        throw new ArgumentException("Unknown argument: " + argument);
                    }
                    else if (oldFile == null) {
                        oldFile = argument;
                    }
                    else if (newFile == null) {
                        newFile = argument;
                    }
                    else {
                        throw new ArgumentException("Specify only two workbook paths.");
                    }
                }
                var config = ExcelTools.Changes.ConfigurationReader.Read(configFile, profileName, oldFile, newFile);
                Console.WriteLine("Configuration: " + config.Path);
                Console.WriteLine("Profile: " + config.Profile);
                Console.WriteLine("Old: " + config.Old);
                Console.WriteLine("New: " + config.New);
                Console.WriteLine("Output: " + config.Output);
                foreach (var comparison in config.Comparisons) {
                    Console.WriteLine(
                        "Sheets: {0} -> {1}; key: {2}; threshold: {3}",
                        comparison.OldSheet,
                        comparison.NewSheet,
                        comparison.Key,
                        comparison.Threshold
                    );
                }
                if (validate) {
                    Console.WriteLine("Configuration and paths validated; workbook contents were not read.");
                    return;
                }
                Console.WriteLine("Reading old workbook...");
                var before = ExcelTools.OpenXml.WorkbookReader.Read(config.Old, config.Comparisons.Select(c => c.OldSheet));
                Console.WriteLine("Reading new workbook...");
                var after = ExcelTools.OpenXml.WorkbookReader.Read(config.New, config.Comparisons.Select(c => c.NewSheet));
                Console.WriteLine("Comparing...");
                var differences = TableComparer.Compare(before, after, config.Comparisons);
                var method = config.Comparisons.Any(comparison => comparison.MatchByRow)
                    ? "В парах с match=\"row\" строки сопоставляются по номеру. Остальные пары сравниваются по ключу и порогу."
                    : null;
                HtmlDiffWriter.Write(config.Output, config.Profile, config.Old, config.New, differences, method);
                Console.WriteLine(
                    "Changed: {0}; added: {1}; deleted: {2}",
                    differences.Sum(s => s.Count("change")),
                    differences.Sum(s => s.Count("add")),
                    differences.Sum(s => s.Count("delete"))
                );
                Console.WriteLine("Created: " + config.Output);
            });
        }

        private static void WritePreparedReport(string path) {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            var report = ExcelTools.Changes.PreparedReport.Read(path);
            XDocument document;
            using (var reader = XmlReader.Create(report.ConfigurationPath, settings)) {
                document = XDocument.Load(reader);
            }
            var profile = document.Root.Elements("profile").Single(element => (string)element.Attribute("name") == report.SyncProfile);
            var section = profile.Elements("externalTools").Elements("tool").Elements("tableDiff").Single();
            var config = ExcelTools.Changes.ConfigurationReader.Read(section, report.ConfigurationPath, report.ReportProfile, false);
            var outputs = ExcelTools.Changes.ConfigurationReader.ReadOutputs(
                profile, report.ConfigurationPath, Path.GetFileNameWithoutExtension(config.Old), "workbook", "previewReport", "appliedReport"
            );
            if (outputs != null) {
                config.Output = outputs[report.StagedWorkbook == null ? "previewReport" : "appliedReport"];
                if (report.StagedWorkbook != null) {
                    config.New = ((bool?)profile.Attribute("overwriteTarget") ?? false) ? config.Old : outputs["workbook"];
                }
            }
            if (!String.Equals(config.Output, report.Output, StringComparison.OrdinalIgnoreCase)
                && !report.Output.StartsWith(config.Output + ".", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Prepared report output does not match the configured report path.");
            }
            var differences = report.Differences;
            if (report.StagedWorkbook != null) {
                var before = ExcelTools.OpenXml.WorkbookReader.Read(config.Old, differences.Select(sheet => sheet.OldSheet));
                var after = ExcelTools.OpenXml.WorkbookReader.Read(report.StagedWorkbook, differences.Select(sheet => sheet.NewSheet));
                differences = differences.Select(sheet => {
                    var difference = TableComparer.CompareByRowNumber(
                        before.FindSheet(sheet.OldSheet),
                        after.FindSheet(sheet.NewSheet),
                        sheet.OldHeader
                    );
                    difference.Key = sheet.Key;
                    difference.Threshold = sheet.Threshold;
                    return difference;
                }).ToList();
            }
            HtmlDiffWriter.Write(report.Output, config.Profile, config.Old, config.New, differences, report.Method);
            Console.WriteLine("Created: " + report.Output);
        }
    }
}