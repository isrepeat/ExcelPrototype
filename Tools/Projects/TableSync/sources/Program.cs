using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelTools.TableSync {
    internal static class Program {
        private sealed class Proposal {
            public ExcelTools.Tables.TableMatch Match;
            public string TargetSheet;
            public string SourceSheet;
            public string Key;
            public string Fallback;
            public ExcelTools.Tables.RowSnapshot TargetHeader;
            public List<ExcelTools.OpenXml.TextCellEdit> Edits = new List<ExcelTools.OpenXml.TextCellEdit>();
        }

        private static int Main(string[] args) {
            return ExcelTools.Tooling.Console.ConsoleTool.Run(args, "TableSync.log", () => {
                var options = args.Where(argument => argument != "--no-pause").ToArray();
                if (options.Contains("--help")) {
                    Console.WriteLine("Usage: TableSync.exe [configuration.xml profile] [--no-pause]");
                    return;
                }
                if (options.Length != 0 && options.Length != 2) {
                    throw new ArgumentException("Specify both configuration.xml and profile, or run without arguments.");
                }
                var configuration = options.Length == 0
                    ? ExcelTools.Tooling.ToolPaths.FindConfiguration(AppDomain.CurrentDomain.BaseDirectory, "TableSync.xml")
                    : ExcelTools.Tooling.ToolPaths.FindConfiguration(Environment.CurrentDirectory, options[0]);
                XDocument document;
                using (var reader = XmlReader.Create(configuration, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit })) {
                    document = XDocument.Load(reader);
                }
                if (document.Root.Name != "tableSync") {
                    throw new InvalidDataException("Expected <tableSync>.");
                }
                var profile = SelectProfile(document.Root, options.Length == 0 ? null : options[1]);
                Run(profile, configuration);
            });
        }

        private static XElement SelectProfile(XElement configuration, string name) {
            var profiles = configuration.Elements("profile").ToDictionary(profile => Required(profile, "name"), StringComparer.Ordinal);
            var defaults = profiles.Values.Where(profile => {
                var attribute = profile.Attribute("default");
                bool isDefault;
                if (attribute == null) {
                    return false;
                }
                if (!Boolean.TryParse(attribute.Value, out isDefault)) {
                    throw new InvalidDataException("Expected true or false for default in profile: " + Required(profile, "name"));
                }
                return isDefault;
            }).ToArray();
            if (defaults.Length > 1) {
                throw new InvalidDataException("Only one profile may have default=\"true\".");
            }
            if (name != null) {
                XElement profile;
                if (!profiles.TryGetValue(name, out profile)) {
                    throw new InvalidDataException("Unknown profile: " + name);
                }
                return profile;
            }
            if (defaults.Length == 1) {
                return defaults[0];
            }
            if (profiles.Count == 1) {
                return profiles.Values.Single();
            }
            throw new InvalidDataException("Set default=\"true\" on one profile in TableSync.xml.");
        }

        private static void Run(XElement profile, string configurationPath) {
            var directory = Path.GetDirectoryName(configurationPath);
            var source = WorkbookPath(profile, "source", directory);
            var target = WorkbookPath(profile, "target", directory);
            var overwrite = (bool?)profile.Attribute("overwriteTarget") ?? false;
            var outputs = ExcelTools.Changes.ConfigurationReader.ReadOutputs(
                profile, configurationPath, Path.GetFileNameWithoutExtension(target), "workbook", "previewReport", "appliedReport"
            );
            var targetStem = Path.Combine(Path.GetDirectoryName(target), Path.GetFileNameWithoutExtension(target));
            var output = overwrite ? target : outputs != null ? outputs["workbook"] : profile.Attribute("output") == null
                ? targetStem + " (synced)" + Path.GetExtension(target) : WorkbookPath(profile, "output", directory);
            if (String.Equals(source, target, StringComparison.OrdinalIgnoreCase)
                || (!overwrite && (String.Equals(source, output, StringComparison.OrdinalIgnoreCase) || File.Exists(output)))) {
                throw new InvalidOperationException("Source and target must be distinct. A copy requires a new output path.");
            }
            if (overwrite && profile.Attribute("output") != null) {
                throw new InvalidDataException("Remove the output attribute when overwriteTarget is true.");
            }
            var mappings = profile.Elements("mapping").ToList();
            var targetSheets = mappings.Select(mapping => Required(mapping, "targetSheet")).ToArray();
            if (mappings.Count == 0 || targetSheets.Distinct(StringComparer.OrdinalIgnoreCase).Count() != mappings.Count) {
                throw new InvalidDataException("Configure one mapping per target sheet.");
            }

            var outputDirectory = Path.GetDirectoryName(output);
            var profileName = Required(profile, "name");
            var tool = profile.Elements("externalTools").Elements("tool")
                .SingleOrDefault(element => element.Element("tableDiff") != null);
            if (tool == null) {
                throw new InvalidDataException("Configure externalTools/tool with an embedded tableDiff section.");
            }
            var executable = ExcelTools.Tooling.ToolPaths.Resolve(directory, Required(tool, "executable"));
            if (!File.Exists(executable) || !String.Equals(Path.GetExtension(executable), ".exe", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("TableDiff executable not found or invalid: " + executable);
            }
            var section = tool.Elements("tableDiff").Single();
            var previewReport = ReadReport(profile, section, configurationPath, "previewDiff", target, source,
                mappings.Select(mapping => Tuple.Create(Required(mapping, "targetSheet"), Required(mapping, "sourceSheet"))));
            var appliedReport = ReadReport(profile, section, configurationPath, "appliedDiff", target, outputs == null ? output : null,
                targetSheets.Select(sheet => Tuple.Create(sheet, sheet)));
            var previewPath = outputs == null ? previewReport.Output : outputs["previewReport"];
            var appliedPath = outputs == null ? appliedReport.Output : outputs["appliedReport"];
            if (String.Equals(previewPath, appliedPath, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Preview and applied reports must have distinct output paths.");
            }
            using (var targetStream = File.Open(target, FileMode.Open, FileAccess.Read,
                overwrite ? FileShare.Read | FileShare.Delete : FileShare.Read)) {
                var targetBook = ExcelTools.OpenXml.WorkbookReader.Read(target, targetSheets);
                ExcelTools.Tables.WorkbookSnapshot sourceBook;
                using (File.Open(source, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    sourceBook = ExcelTools.OpenXml.WorkbookReader.Read(source, mappings.Select(mapping => Required(mapping, "sourceSheet")));
                }
                var proposals = mappings.SelectMany(mapping => Preview(mapping, targetBook, sourceBook)).ToList();
                RunTableDiff(executable, new ExcelTools.Changes.PreparedReport {
                    ConfigurationPath = configurationPath, SyncProfile = profileName, ReportProfile = previewReport.Profile,
                    Output = previewPath, Differences = ArrangeReport(PreviewDiff(proposals), previewReport),
                    Method = "В отчёте показаны предлагаемые правки целевой книги. Подтверждение в консоли применяет весь список."
                });
                OpenReport(previewPath);
                if (proposals.Count == 0) {
                    Console.WriteLine("No changes available for approval.");
                    return;
                }
                for (var index = 0; index < proposals.Count; index++) {
                    ShowProposal(index + 1, proposals[index]);
                }

                Console.WriteLine("Output: " + output);
                Console.Write("Apply all {0} proposals? Type \"y\" or \"yes\" to confirm, or press Enter to cancel: ", proposals.Count);
                var response = Console.ReadLine();
                if (!String.Equals(response, "y", StringComparison.OrdinalIgnoreCase)
                    && !String.Equals(response, "yes", StringComparison.OrdinalIgnoreCase)) {
                    Console.WriteLine("Canceled.");
                    return;
                }
                var edits = proposals.SelectMany(proposal => proposal.Edits).ToList();
                Directory.CreateDirectory(outputDirectory);
                var staged = ExcelTools.Tooling.OutputPublisher.CreateTemporaryPath(output, ".xlsx");
                var stagedReport = ExcelTools.Tooling.OutputPublisher.CreateTemporaryPath(appliedPath, ".html");
                try {
                    using (var stream = File.Open(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                        targetStream.CopyTo(stream);
                    }
                    ExcelTools.OpenXml.WorkbookWriter.ApplyTextEdits(staged, edits);
                    RunTableDiff(executable, new ExcelTools.Changes.PreparedReport {
                        ConfigurationPath = configurationPath, SyncProfile = profileName, ReportProfile = appliedReport.Profile,
                        Output = stagedReport, StagedWorkbook = staged,
                        Differences = ArrangeReport(AppliedSheets(proposals), appliedReport),
                        Method = "Фактические изменения целевой книги. Строки сопоставлены по номеру, включая строки с изменённым ключом."
                    });
                    if (overwrite) {
                        File.Replace(staged, output, null);
                    }
                    else {
                        File.Move(staged, output);
                    }
                    ExcelTools.Tooling.OutputPublisher.Publish(stagedReport, appliedPath);
                }
                finally {
                    if (File.Exists(staged)) {
                        File.Delete(staged);
                    }
                    if (File.Exists(stagedReport)) {
                        File.Delete(stagedReport);
                    }
                }
                Console.WriteLine("Saved " + edits.Count + " approved cell edits: " + output);
                OpenReport(appliedPath);
            }
        }

        private static ExcelTools.Changes.Configuration ReadReport(
            XElement profile, XElement section, string configurationPath, string attribute, string oldFile, string newFile,
            IEnumerable<Tuple<string, string>> sheetPairs) {
            var report = ExcelTools.Changes.ConfigurationReader.Read(section, configurationPath, Required(profile, attribute), false);
            if (!String.Equals(report.Old, oldFile, StringComparison.OrdinalIgnoreCase)
                || (newFile != null && !String.Equals(report.New, newFile, StringComparison.OrdinalIgnoreCase))) {
                throw new InvalidDataException("Workbook paths in TableDiff profile '" + report.Profile + "' must match the synchronization paths.");
            }
            var configured = report.Comparisons.Select(pair => Tuple.Create(pair.OldSheet, pair.NewSheet));
            if (!new HashSet<Tuple<string, string>>(sheetPairs).SetEquals(configured)) {
                throw new InvalidDataException("Sheet pairs in TableDiff profile '" + report.Profile + "' must match the synchronization mappings.");
            }
            return report;
        }

        private static List<ExcelTools.Changes.SheetDiff> ArrangeReport(List<ExcelTools.Changes.SheetDiff> differences, ExcelTools.Changes.Configuration report) {
            var byPair = differences.ToDictionary(diff => Tuple.Create(diff.OldSheet, diff.NewSheet));
            var arranged = new List<ExcelTools.Changes.SheetDiff>();
            foreach (var comparison in report.Comparisons) {
                ExcelTools.Changes.SheetDiff difference;
                if (byPair.TryGetValue(Tuple.Create(comparison.OldSheet, comparison.NewSheet), out difference)) {
                    difference.Key = comparison.Key;
                    difference.Threshold = comparison.Threshold;
                    arranged.Add(difference);
                }
            }
            return arranged;
        }

        private static void RunTableDiff(string executable, ExcelTools.Changes.PreparedReport report) {
            var requestPath = Path.Combine(Path.GetTempPath(), "TableSync-" + Guid.NewGuid().ToString("N") + ".xml");
            try {
                using (var stream = File.Open(requestPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    ExcelTools.Changes.PreparedReport.Write(stream, report);
                }
                using (var process = Process.Start(new ProcessStartInfo(executable, "--prepared-report \"" + requestPath + "\" --no-pause") {
                    UseShellExecute = false,
                    WorkingDirectory = Path.GetDirectoryName(executable)
                })) {
                    process.WaitForExit();
                    if (process.ExitCode != 0 || !File.Exists(report.Output)) {
                        throw new InvalidOperationException("TableDiff failed to create the report. Exit code: " + process.ExitCode);
                    }
                }
            }
            finally {
                File.Delete(requestPath);
            }
        }

        private static List<ExcelTools.Changes.SheetDiff> PreviewDiff(List<Proposal> proposals) {
            return proposals.Select((proposal, index) => new { Proposal = proposal, Number = index + 1 })
                .GroupBy(item => item.Proposal.TargetSheet, StringComparer.Ordinal).Select(group => {
                    var first = group.First().Proposal;
                    var columns = group.SelectMany(item => item.Proposal.Edits).Select(edit => edit.Column - 1)
                        .Distinct().OrderBy(column => column).ToArray();
                    var header = new ExcelTools.Tables.RowSnapshot {
                        Number = first.TargetHeader.Number,
                        Values = new[] { "№ предложения" }.Concat(columns.Select(first.TargetHeader.ValueAt)).ToArray()
                    };
                    var addresses = new[] { -1 }.Concat(columns).ToArray();
                    var difference = new ExcelTools.Changes.SheetDiff {
                        Name = group.Key, OldSheet = group.Key, NewSheet = first.SourceSheet,
                        Width = header.Values.Length, OldHeader = header, NewHeader = header,
                        OldColumns = addresses, NewColumns = addresses
                    };
                    foreach (var item in group) {
                        var proposal = item.Proposal;
                        var row = proposal.Match.Target;
                        var edits = proposal.Edits.ToDictionary(edit => edit.Column - 1, edit => edit.Value);
                        var before = new[] { item.Number.ToString() }.Concat(columns.Select(row.ValueAt)).ToArray();
                        var after = (string[])before.Clone();
                        for (var index = 0; index < columns.Length; index++) {
                            string value;
                            if (edits.TryGetValue(columns[index], out value)) {
                                after[index + 1] = value;
                            }
                        }
                        difference.Changes.Add(new ExcelTools.Changes.RowChange {
                            Kind = "change", Before = new ExcelTools.Tables.RowSnapshot { Number = row.Number, Values = before },
                            After = new ExcelTools.Tables.RowSnapshot { Number = row.Number, Values = after },
                            Ambiguous = proposal.Match.ByFallback,
                            ReviewNote = proposal.Match.ByFallback ? "Сопоставлено по резервному полю. Проверьте пару перед подтверждением." : null
                        });
                    }
                    return difference;
                }).ToList();
        }

        private static List<ExcelTools.Changes.SheetDiff> AppliedSheets(List<Proposal> proposals) {
            return proposals.GroupBy(proposal => proposal.TargetSheet, StringComparer.Ordinal)
                .Select(group => new ExcelTools.Changes.SheetDiff {
                    OldSheet = group.Key, NewSheet = group.Key, OldHeader = group.First().TargetHeader
                }).ToList();
        }

        private static void OpenReport(string path) {
            Console.WriteLine("Report: " + path);
            if (!Console.IsInputRedirected) {
                try {
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
                catch (Exception error) {
                    Console.Error.WriteLine("Open the report manually: " + error.Message);
                }
            }
        }

        private static IEnumerable<Proposal> Preview(XElement mapping, ExcelTools.Tables.WorkbookSnapshot targetBook, ExcelTools.Tables.WorkbookSnapshot sourceBook) {
            var target = targetBook.FindSheet(Required(mapping, "targetSheet"));
            var source = sourceBook.FindSheet(Required(mapping, "sourceSheet"));
            var targetKeyName = Required(mapping, "targetKey");
            var sourceKeyName = Required(mapping, "sourceKey");
            var targetFallback = (string)mapping.Attribute("targetFallback");
            var sourceFallback = (string)mapping.Attribute("sourceFallback");
            if ((targetFallback == null) != (sourceFallback == null)) {
                throw new InvalidDataException("Configure both fallback columns or omit both.");
            }
            var fields = mapping.Elements("field").Select(field => new {
                Target = Required(field, "target"),
                Source = Required(field, "source"),
                IgnorePattern = field.Attribute("compareIgnorePattern") == null ? null : new Regex(
                    Required(field, "compareIgnorePattern"),
                    RegexOptions.CultureInvariant,
                    TimeSpan.FromSeconds(1)
                )
            }).ToList();
            if (fields.Count == 0 || fields.Select(field => field.Target).Distinct(StringComparer.Ordinal).Count() != fields.Count
                || fields.Any(field => field.Target == targetKeyName && field.Source != sourceKeyName)) {
                throw new InvalidDataException("Configure unique target fields. A key update must come from the source key.");
            }
            if (targetFallback != null && (!fields.Any(field => field.Target == targetKeyName)
                || targetFallback == targetKeyName || sourceFallback == sourceKeyName)) {
                throw new InvalidDataException("Fallback matching requires distinct fallback columns and a key field mapping.");
            }
            var targetNames = fields.Select(field => field.Target).Concat(new[] { targetKeyName, targetFallback }).Where(name => name != null);
            var sourceNames = fields.Select(field => field.Source).Concat(new[] { sourceKeyName, sourceFallback }).Where(name => name != null);
            var targetHeader = ExcelTools.Tables.TableMatching.FindHeader(target, targetNames, (int?)mapping.Attribute("targetHeaderRow"));
            var sourceHeader = ExcelTools.Tables.TableMatching.FindHeader(source, sourceNames, (int?)mapping.Attribute("sourceHeaderRow"));
            var targetKey = ExcelTools.Tables.TableMatching.FindColumn(targetHeader, targetKeyName);
            var sourceKey = ExcelTools.Tables.TableMatching.FindColumn(sourceHeader, sourceKeyName);
            var columns = fields.Select(field => new {
                Target = ExcelTools.Tables.TableMatching.FindColumn(targetHeader, field.Target),
                Source = ExcelTools.Tables.TableMatching.FindColumn(sourceHeader, field.Source),
                field.IgnorePattern
            }).ToArray();
            var targetRows = ExcelTools.Tables.TableMatching.DataRows(target, targetHeader, (int?)mapping.Attribute("targetLastRow"));
            var sourceRows = ExcelTools.Tables.TableMatching.DataRows(source, sourceHeader, (int?)mapping.Attribute("sourceLastRow"));
            var targetFallbackColumn = targetFallback == null ? (int?)null : ExcelTools.Tables.TableMatching.FindColumn(targetHeader, targetFallback);
            var matches = ExcelTools.Tables.TableMatching.Match(
                targetRows,
                sourceRows,
                targetKey,
                sourceKey,
                targetFallbackColumn,
                sourceFallback == null ? (int?)null : ExcelTools.Tables.TableMatching.FindColumn(sourceHeader, sourceFallback)
            );

            foreach (var row in targetRows) {
                ExcelTools.Tables.TableMatch match;
                if (!matches.TryGetValue(row.Number, out match)) {
                    Console.WriteLine("Review manually: {0}!{1}, key {2}", target.Name, row.Number, Display(row.ValueAt(targetKey)));
                    continue;
                }
                var proposal = new Proposal {
                    Match = match,
                    TargetSheet = target.Name,
                    SourceSheet = source.Name,
                    Key = row.ValueAt(targetKey),
                    TargetHeader = targetHeader,
                    Fallback = match.ByFallback ? row.ValueAt(targetFallbackColumn.Value) : null
                };
                foreach (var column in columns) {
                    var oldValue = row.ValueAt(column.Target);
                    var newValue = match.Source.ValueAt(column.Source);
                    if (oldValue == newValue || (column.IgnorePattern != null
                        && column.IgnorePattern.Replace(oldValue, "") == column.IgnorePattern.Replace(newValue, ""))) {
                        continue;
                    }
                    if (String.IsNullOrWhiteSpace(newValue) || newValue.StartsWith("=", StringComparison.Ordinal)
                        || oldValue.StartsWith("=", StringComparison.Ordinal)) {
                        Console.WriteLine("Skipped blank or formula change: {0}!{1}{2}",
                            target.Name, ExcelTools.Tables.CellAddress.ColumnName(column.Target + 1), row.Number);
                        continue;
                    }
                    proposal.Edits.Add(new ExcelTools.OpenXml.TextCellEdit {
                        Sheet = target.Name, Row = row.Number, Column = column.Target + 1, Value = newValue
                    });
                }
                if (proposal.Edits.Count != 0) {
                    yield return proposal;
                }
            }
        }

        private static void ShowProposal(int number, Proposal proposal) {
            var row = proposal.Match.Target;
            Console.WriteLine("[{0}] {1}!{2} <- {3}!{4} | key {5}{6}", number, proposal.TargetSheet, row.Number,
                proposal.SourceSheet, proposal.Match.Source.Number, Display(proposal.Key),
                proposal.Match.ByFallback ? " | VERIFY fallback match " + Display(proposal.Fallback) : "");
            foreach (var edit in proposal.Edits) {
                Console.WriteLine("  {0} ({1}{2}): {3} -> {4}", proposal.TargetHeader.ValueAt(edit.Column - 1),
                    ExcelTools.Tables.CellAddress.ColumnName(edit.Column), edit.Row, Display(row.ValueAt(edit.Column - 1)), Display(edit.Value));
            }
        }

        private static string Display(string value) {
            return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t") + "\"";
        }

        private static string WorkbookPath(XElement profile, string attribute, string directory) {
            var path = ExcelTools.Tooling.ToolPaths.Resolve(directory, Required(profile, attribute));
            if (!String.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Expected an .xlsx path for " + attribute);
            }
            return path;
        }

        private static string Required(XElement element, string attribute) {
            var value = (string)element.Attribute(attribute);
            if (String.IsNullOrWhiteSpace(value)) {
                throw new InvalidDataException("Missing " + element.Name + "/@" + attribute);
            }
            return value;
        }
    }
}