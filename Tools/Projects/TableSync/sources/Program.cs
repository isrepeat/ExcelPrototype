using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;

namespace ExcelTools.TableSync {
    internal static class Program {
        private sealed class Table {
            public ExcelTools.Tables.SheetSnapshot Sheet;
            public ExcelTools.Tables.RowSnapshot Header;
            public List<ExcelTools.Tables.RowSnapshot> Rows;
            public int LastRow;
            public int Key;
            public string TableName;
            public int FirstColumn;
            public int LastColumn;
        }

        private sealed class Update {
            public string Book;
            public Table Table;
            public bool All;
            public List<Field> Fields;
        }

        private sealed class Field {
            public int Source;
            public int Target;
            public Regex Ignore;
        }

        private static int Main(string[] args) {
            return ExcelTools.Tooling.Console.ConsoleTool.Run(args, "TableSync.log", () => {
                var options = args.Where(argument => argument != "--no-pause").ToArray();
                if (options.Length > 2) {
                    throw new ArgumentException("Usage: TableSync.exe [configuration.xml [profile]] [--no-pause]");
                }
                var directory = options.Length == 0 ? AppDomain.CurrentDomain.BaseDirectory : Environment.CurrentDirectory;
                var config = ExcelTools.Tooling.ToolPaths.FindConfiguration(directory, options.FirstOrDefault() ?? "TableSync.xml");
                using (var reader = XmlReader.Create(config, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit })) {
                    Run(XDocument.Load(reader).Root, config, options.Length == 2 ? options[1] : null);
                }
            });
        }

        private static void Run(XElement config, string configuration, string profileName) {
            if (config == null || config.Name != "tableSync") {
                throw new InvalidDataException("Expected tableSync configuration.");
            }
            var profiles = config.Elements("profile").ToDictionary(element => Required(element, "name"), StringComparer.Ordinal);
            var preferred = profiles.Values.Where(element => (bool?)element.Attribute("default") == true).SingleOrDefault();
            var profile = profileName == null ? preferred ?? profiles.Values.Single() : profiles[profileName];
            var directory = Path.GetDirectoryName(configuration);
            var books = profile.Elements("books").Elements("book").ToDictionary(
                book => Required(book, "name"), book => Resolve(book, directory, "file"), StringComparer.Ordinal);
            var plugin = profile.Elements("plugin").SingleOrDefault();
            if (plugin != null && profile.Elements().Any(element => element.Name == "source" || element.Name == "update" || element.Name == "match")) {
                throw new InvalidDataException("Use either a plugin or the built-in rules.");
            }
            var source = plugin == null ? profile.Elements("source").Single() : null;
            var sourceBook = source == null ? null : Required(source, "book");
            var outputs = ExcelTools.Changes.ConfigurationReader.ReadOutputs(profile, configuration,
                sourceBook == null ? Required(profile, "name") : Path.GetFileNameWithoutExtension(books[sourceBook]), "previewReport", "appliedReport");
            if (outputs == null || String.Equals(outputs["previewReport"], outputs["appliedReport"], StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Configure distinct previewReport and appliedReport outputs.");
            }
            var executable = Resolve(profile, directory, "tableDiff");
            var actions = profile.Elements("update").ToList();
            if (profile.Element("recordKeyChange") != null || (plugin == null && actions.Count == 0)
                || books.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != books.Count) {
                throw new InvalidDataException("Configure a plugin or update rules and distinct workbook paths. History rules belong in plugins.");
            }

            var locks = new List<FileStream>();
            try {
                foreach (var path in books.Values) {
                    if (!String.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                        throw new InvalidDataException("Expected an .xlsx workbook: " + path);
                    }
                    locks.Add(File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read));
                }
                var tables = new List<Tuple<string, Table>>();
                var edits = new Dictionary<string, List<ExcelTools.OpenXml.TextCellEdit>>(StringComparer.Ordinal);
                var conflicts = new List<string>();
                if (plugin == null) {
                    var sourceTable = ReadTable(books[sourceBook], Required(source, "sheet"), (string)source.Attribute("table"),
                        Required(source, "key"));
                    var updates = actions.Select(action => ReadUpdate(action, books, sourceBook, sourceTable)).ToList();
                    if (updates.GroupBy(update => Tuple.Create(update.Book, update.Table.Sheet.Name)).Any(group => group.Count() != 1)) {
                        throw new InvalidDataException("Configure one update per destination sheet.");
                    }
                    foreach (var update in updates) {
                        tables.Add(Tuple.Create(update.Book, update.Table));
                        if (!edits.ContainsKey(update.Book)) {
                            edits.Add(update.Book, new List<ExcelTools.OpenXml.TextCellEdit>());
                        }
                    }
                    Plan(profile, sourceTable, updates, edits);
                }
                else {
                    sourceBook = PlanPlugin(plugin, directory, books, tables, edits, conflicts);
                }
                if (edits.Values.All(items => items.Count == 0)) {
                    System.Console.WriteLine("No changes available for approval.");
                    return;
                }
                var preview = outputs["previewReport"];
                var reportProfile = Required(profile, "name");
                var primary = edits.First(pair => pair.Value.Count > 0).Key;
                WriteReport(executable, preview, reportProfile, books[primary], books[sourceBook],
                    tables.Select(item => Diff(item.Item1, item.Item2, edits[item.Item1], null)), conflicts);
                OpenReport(preview);
                System.Console.Write("Apply all {0} cell edits? Type \"y\" or \"yes\" to confirm: ", edits.Values.Sum(items => items.Count));
                var answer = System.Console.ReadLine();
                if (!String.Equals(answer, "y", StringComparison.OrdinalIgnoreCase)
                    && !String.Equals(answer, "yes", StringComparison.OrdinalIgnoreCase)) {
                    System.Console.WriteLine("Canceled.");
                    return;
                }

                var result = Path.Combine(Path.GetDirectoryName(outputs["appliedReport"]), "TableSync "
                    + DateTime.Now.ToString("HHmmssfff", CultureInfo.InvariantCulture));
                var stage = result + "." + Guid.NewGuid().ToString("N") + ".tmp";
                Directory.CreateDirectory(stage);
                try {
                    var copies = edits.Keys.ToDictionary(book => book, book => Path.Combine(stage, SyncedName(books[book])));
                    if (copies.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != copies.Count) {
                        throw new InvalidDataException("Destination workbook file names must be distinct.");
                    }
                    foreach (var copy in copies) {
                        File.Copy(books[copy.Key], copy.Value);
                        foreach (var item in tables.Where(item => item.Item1 == copy.Key)) {
                            var changes = edits[copy.Key].Where(edit => edit.Sheet == item.Item2.Sheet.Name).ToList();
                            if (changes.Count > 0) {
                                ExcelTools.OpenXml.WorkbookWriter.ApplyTextEdits(copy.Value, changes,
                                    changes.Any(edit => edit.Row > item.Item2.LastRow) ? item.Item2.TableName : null);
                            }
                        }
                    }
                    var actual = tables.Select(item => Diff(item.Item1, item.Item2, edits[item.Item1],
                        ExcelTools.OpenXml.WorkbookReader.Read(copies[item.Item1], new[] { item.Item2.Sheet.Name })
                            .FindSheet(item.Item2.Sheet.Name))).ToList();
                    var reportName = Path.GetFileName(outputs["appliedReport"]);
                    WriteReport(executable, Path.Combine(stage, reportName), reportProfile, books[primary], copies[primary], actual);
                    Directory.Move(stage, result);
                    OpenReport(Path.Combine(result, reportName));
                }
                finally {
                    if (Directory.Exists(stage)) {
                        foreach (var file in Directory.GetFiles(stage)) {
                            File.Delete(file);
                        }
                        Directory.Delete(stage);
                    }
                }
            }
            finally {
                foreach (var stream in locks) {
                    stream.Dispose();
                }
            }
        }

        private static string PlanPlugin(
            XElement config, string directory, Dictionary<string, string> books, List<Tuple<string, Table>> tables,
            Dictionary<string, List<ExcelTools.OpenXml.TextCellEdit>> edits, List<string> conflicts) {
            var assemblyPath = Resolve(config, directory, "assembly");
            var candidates = Assembly.LoadFrom(assemblyPath).GetExportedTypes()
                .Where(candidate => candidate.IsClass
                    && !candidate.IsAbstract
                    && !candidate.ContainsGenericParameters
                    && typeof(ExcelTools.TableSync.Contracts.ISyncPlugin).IsAssignableFrom(candidate))
                .ToArray();
            if (candidates.Length != 1) {
                throw new InvalidDataException("Expected exactly one public concrete ISyncPlugin implementation in "
                    + assemblyPath + ". Found: " + candidates.Length + ".");
            }
            var type = candidates[0];
            if (type.GetConstructor(Type.EmptyTypes) == null) {
                throw new InvalidDataException("Plugin requires a public parameterless constructor: " + type.FullName);
            }
            var plugin = (ExcelTools.TableSync.Contracts.ISyncPlugin)Activator.CreateInstance(type);
            var snapshots = new Dictionary<ExcelTools.TableSync.Contracts.SyncTable, Table>();
            var context = new ExcelTools.TableSync.Contracts.SyncContext((book, sheet, name, key) => {
                var table = ReadTable(books[book], sheet, name, key);
                var snapshot = new ExcelTools.TableSync.Contracts.SyncTable(book, sheet, name, table.LastRow, table.Header.Values,
                    table.Rows.Select(row => new ExcelTools.TableSync.Contracts.SyncRow(row.Number, row.Values)));
                snapshots.Add(snapshot, table);
                return snapshot;
            });
            var plan = plugin.BuildPlan(context);
            if (plan == null || String.IsNullOrWhiteSpace(plan.SourceBook) || !books.ContainsKey(plan.SourceBook)) {
                throw new InvalidDataException("Plugin must return a plan with a configured source book.");
            }
            var targets = new Dictionary<Tuple<string, string>, ExcelTools.TableSync.Contracts.SyncTable>();
            var cells = new HashSet<Tuple<string, string, int, int>>();
            foreach (var group in plan.Groups) {
                if (!String.IsNullOrWhiteSpace(group.Conflict)) {
                    var message = group.Description + ": " + group.Conflict;
                    conflicts.Add(message);
                    System.Console.WriteLine("Manual review: " + message);
                    continue;
                }
                foreach (var edit in group.Edits) {
                    var snapshot = edit.Table;
                    Table table;
                    if (!snapshots.TryGetValue(snapshot, out table) || snapshot.Book == plan.SourceBook
                        || edit.Column < table.FirstColumn || edit.Column > table.LastColumn || edit.Row <= table.Header.Number
                        || edit.Row > 1048576 || edit.Value == null) {
                        throw new InvalidDataException("Invalid edit in plugin plan: " + group.Description);
                    }
                    var destination = Tuple.Create(snapshot.Book, snapshot.Sheet);
                    ExcelTools.TableSync.Contracts.SyncTable previous;
                    if (!targets.TryGetValue(destination, out previous)) {
                        targets.Add(destination, snapshot);
                        tables.Add(Tuple.Create(snapshot.Book, table));
                    }
                    else if (!Object.ReferenceEquals(previous, snapshot)) {
                        throw new InvalidDataException("Use one snapshot per destination sheet.");
                    }
                    if (!cells.Add(Tuple.Create(snapshot.Book, snapshot.Sheet, edit.Row, edit.Column))) {
                        throw new InvalidDataException("Duplicate cell edit in plugin plan: " + group.Description);
                    }
                    List<ExcelTools.OpenXml.TextCellEdit> bookEdits;
                    if (!edits.TryGetValue(snapshot.Book, out bookEdits)) {
                        bookEdits = new List<ExcelTools.OpenXml.TextCellEdit>();
                        edits.Add(snapshot.Book, bookEdits);
                    }
                    bookEdits.Add(new ExcelTools.OpenXml.TextCellEdit {
                        Sheet = snapshot.Sheet, Row = edit.Row, Column = edit.Column + 1, Value = edit.Value
                    });
                }
            }
            foreach (var target in targets.Values) {
                var existing = target.Rows.ToDictionary(row => row.Number);
                var changes = edits[target.Book].Where(edit => edit.Sheet == target.Sheet).ToList();
                var rows = changes.Select(edit => edit.Row).Distinct().OrderBy(row => row);
                var next = target.LastRow + 1;
                foreach (var row in rows) {
                    if (!existing.ContainsKey(row) && (target.Table == null || row != next++)) {
                        throw new InvalidDataException("Plugin append must extend a named table with consecutive rows.");
                    }
                }
                if (changes.Any(edit => existing.ContainsKey(edit.Row)
                    && existing[edit.Row].ValueAt(edit.Column - 1).StartsWith("=", StringComparison.Ordinal))) {
                    throw new InvalidDataException("Plugin plan cannot replace formulas.");
                }
            }
            return plan.SourceBook;
        }

        private static Update ReadUpdate(XElement action, Dictionary<string, string> books, string sourceBook, Table source) {
            var book = Required(action, "book");
            var rows = (string)action.Attribute("rows") ?? "unique";
            if (book == sourceBook || (rows != "all" && rows != "unique")) {
                throw new InvalidDataException("An update requires a destination book and rows=\"all\" or rows=\"unique\".");
            }
            var table = ReadTable(books[book], Required(action, "sheet"), (string)action.Attribute("table"), Required(action, "key"));
            var fields = action.Elements("field").Select(field => new Field {
                Source = Column(source, Required(field, "source")),
                Target = Column(table, Required(field, "target")),
                Ignore = Pattern(field, "compareIgnorePattern")
            }).ToList();
            if (fields.Count == 0 || fields.Select(field => field.Target).Distinct().Count() != fields.Count
                || fields.Any(field => field.Target == table.Key && (field.Source != source.Key || field.Ignore != null))) {
                throw new InvalidDataException("Configure distinct destination fields. A key may only be assigned from the source key.");
            }
            return new Update { Book = book, Table = table, All = rows == "all", Fields = fields };
        }

        private static void Plan(
            XElement profile, Table source, List<Update> updates,
            Dictionary<string, List<ExcelTools.OpenXml.TextCellEdit>> edits) {
            var fallback = profile.Elements("match").Elements("fallback").SingleOrDefault();
            var sourceKeys = source.Rows.Where(row => Valid(row.ValueAt(source.Key))).ToLookup(row => row.ValueAt(source.Key), StringComparer.Ordinal);
            var groups = updates.SelectMany(update => update.Table.Rows.Select(row => new { Update = update, Row = row }))
                .Where(item => Valid(item.Row.ValueAt(item.Update.Table.Key)))
                .ToLookup(item => item.Row.ValueAt(item.Update.Table.Key), StringComparer.Ordinal);
            var namesByKey = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
            ILookup<string, ExcelTools.Tables.RowSnapshot> sourceNames = null;
            Regex pattern = null;
            var sourceName = -1;
            if (fallback != null) {
                if (Required(fallback, "by") != "uniqueName") {
                    throw new InvalidDataException("Expected fallback by=\"uniqueName\".");
                }
                sourceName = Column(source, Required(fallback, "sourceColumn"));
                pattern = Pattern(fallback, "compareIgnorePattern");
                sourceNames = source.Rows.ToLookup(row => Normalize(row.ValueAt(sourceName), pattern), StringComparer.Ordinal);
                foreach (var search in fallback.Elements("search")) {
                    var update = updates.Single(item => item.Book == Required(search, "book")
                        && item.Table.Sheet.Name == Required(search, "sheet"));
                    if ((string)search.Attribute("table") != update.Table.TableName
                        || Column(update.Table, Required(search, "key")) != update.Table.Key) {
                        throw new InvalidDataException("A fallback search must identify the same range and key as an update.");
                    }
                    var column = Column(update.Table, Required(search, "nameColumn"));
                    foreach (var row in update.Table.Rows) {
                        var key = row.ValueAt(update.Table.Key);
                        var name = Normalize(row.ValueAt(column), pattern);
                        if (!Valid(key) || !Valid(name)) {
                            continue;
                        }
                        HashSet<string> names;
                        if (!namesByKey.TryGetValue(key, out names)) {
                            names = new HashSet<string>(StringComparer.Ordinal);
                            namesByKey.Add(key, names);
                        }
                        names.Add(name);
                    }
                }
            }
            var keysByName = namesByKey.SelectMany(pair => pair.Value.Select(name => Tuple.Create(name, pair.Key)))
                .ToLookup(item => item.Item1, StringComparer.Ordinal);
            foreach (var group in groups) {
                var person = sourceKeys[group.Key].SingleOrDefaultIfUnique();
                var keyChanged = !sourceKeys.Contains(group.Key);
                if (keyChanged && sourceNames != null) {
                    HashSet<string> names;
                    if (namesByKey.TryGetValue(group.Key, out names) && names.Count == 1 && keysByName[names.Single()].Count() == 1) {
                        person = sourceNames[names.Single()].SingleOrDefaultIfUnique();
                    }
                }
                if (person == null) {
                    continue;
                }
                var newKey = person.ValueAt(source.Key);
                if (!Valid(newKey) || sourceKeys[newKey].Count() != 1 || (keyChanged && groups.Contains(newKey))) {
                    continue;
                }

                var pending = new List<Tuple<string, ExcelTools.OpenXml.TextCellEdit>>();
                foreach (var destination in group.GroupBy(item => item.Update)) {
                    if (!destination.Key.All && destination.Count() != 1) {
                        continue;
                    }
                    foreach (var item in destination) {
                        foreach (var field in item.Update.Fields) {
                            var before = item.Row.ValueAt(field.Target);
                            var after = person.ValueAt(field.Source);
                            if (Valid(after) && !before.StartsWith("=", StringComparison.Ordinal)
                                && CompareValue(before, field.Ignore) != CompareValue(after, field.Ignore)) {
                                pending.Add(Tuple.Create(item.Update.Book, new ExcelTools.OpenXml.TextCellEdit {
                                    Sheet = item.Update.Table.Sheet.Name, Row = item.Row.Number, Column = field.Target + 1, Value = after
                                }));
                            }
                        }
                    }
                }
                if (keyChanged && pending.Count > 0) {
                    if (group.GroupBy(item => item.Update).Any(destination =>
                        destination.Key.Fields.Any(field => field.Target == destination.Key.Table.Key)
                        && (!destination.Key.All && destination.Count() != 1))) {
                        System.Console.WriteLine("Manual review: destination key is repeated under rows=\"unique\": " + group.Key);
                        continue;
                    }
                }
                foreach (var item in pending) {
                    edits[item.Item1].Add(item.Item2);
                }
            }
        }

        private static Regex Pattern(XElement element, string attribute) {
            return element.Attribute(attribute) == null ? null
                : new Regex(Required(element, attribute), RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        }

        private static bool Valid(string value) {
            return !String.IsNullOrWhiteSpace(value) && !value.StartsWith("=", StringComparison.Ordinal);
        }

        private static string CompareValue(string value, Regex pattern) {
            return pattern == null ? value : pattern.Replace(value, "");
        }

        private static string Normalize(string value, Regex pattern) {
            return Valid(value) ? CompareValue(value, pattern).Trim() : "";
        }

        private static T SingleOrDefaultIfUnique<T>(this IEnumerable<T> items) where T : class {
            var candidates = items.Take(2).ToArray();
            return candidates.Length == 1 ? candidates[0] : null;
        }

        private static int Column(Table table, string name) {
            return ExcelTools.Tables.TableMatching.FindColumn(table.Header, name);
        }

        private static Table ReadTable(string path, string sheet, string tableName, string key) {
            var snapshot = ExcelTools.OpenXml.WorkbookReader.Read(path, new[] { sheet }).FindSheet(sheet);
            var range = tableName == null ? null : ExcelTools.OpenXml.WorkbookReader.ReadTableRange(path, sheet, tableName).Split(':');
            int? headerRow = range == null ? (int?)null : ExcelTools.Tables.CellAddress.RowNumber(range[0]);
            int? lastRow = range == null ? (int?)null : ExcelTools.Tables.CellAddress.RowNumber(range[1]);
            var header = ExcelTools.Tables.TableMatching.FindHeader(snapshot, new[] { key }, headerRow);
            var table = new Table {
                Sheet = snapshot, Header = header, LastRow = lastRow ?? snapshot.Rows.Max(row => row.Number),
                Rows = ExcelTools.Tables.TableMatching.DataRows(snapshot, header, lastRow),
                FirstColumn = range == null ? 0 : ExcelTools.Tables.CellAddress.ColumnNumber(range[0]) - 1,
                LastColumn = range == null ? header.Values.Length - 1 : ExcelTools.Tables.CellAddress.ColumnNumber(range[1]) - 1
            };
            table.Key = Column(table, key);
            table.TableName = tableName;
            return table;
        }

        private static ExcelTools.Changes.SheetDiff Diff(
            string book, Table table, IEnumerable<ExcelTools.OpenXml.TextCellEdit> edits, ExcelTools.Tables.SheetSnapshot actual) {
            var byNumber = table.Rows.ToDictionary(row => row.Number);
            var saved = actual == null ? null : actual.Rows.ToDictionary(row => row.Number);
            var changes = new ExcelTools.Changes.SheetDiff {
                Name = book + " · " + table.Sheet.Name, OldSheet = table.Sheet.Name, NewSheet = table.Sheet.Name,
                OldHeader = table.Header, NewHeader = table.Header, Width = table.Header.Values.Length,
                OldCount = table.Rows.Count, NewCount = table.Rows.Count
            };
            foreach (var group in edits.Where(edit => edit.Sheet == table.Sheet.Name).GroupBy(edit => edit.Row)) {
                ExcelTools.Tables.RowSnapshot before;
                byNumber.TryGetValue(group.Key, out before);
                var values = Enumerable.Range(0, changes.Width).Select(column => before == null ? "" : before.ValueAt(column)).ToArray();
                foreach (var edit in group) {
                    values[edit.Column - 1] = edit.Value;
                    if (saved != null && saved[group.Key].ValueAt(edit.Column - 1) != edit.Value) {
                        throw new InvalidDataException("Saved value differs from plan: " + table.Sheet.Name + " row " + group.Key);
                    }
                }
                var after = saved == null ? new ExcelTools.Tables.RowSnapshot { Number = group.Key, Values = values } : saved[group.Key];
                changes.Changes.Add(new ExcelTools.Changes.RowChange { Kind = before == null ? "add" : "change", Before = before, After = after });
                changes.NewCount += before == null ? 1 : 0;
            }
            return changes;
        }

        private static void WriteReport(
            string executable, string path, string profile, string before, string after,
            IEnumerable<ExcelTools.Changes.SheetDiff> differences, IEnumerable<string> conflicts = null) {
            new ExcelTools.Changes.PreparedReport {
                Output = path, ReportProfile = profile, OldFile = before, NewFile = after,
                Method = "Отчёт включает изменения по всем книгам профиля."
                    + (conflicts == null ? "" : " " + String.Join(" · ", conflicts)), Differences = differences.ToList()
            }.Run(executable);
        }

        private static void OpenReport(string path) {
            System.Console.WriteLine("Report: " + path);
            if (!System.Console.IsInputRedirected) {
                try {
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
                catch (Exception error) {
                    System.Console.Error.WriteLine(error.Message);
                }
            }
        }

        private static string SyncedName(string path) {
            return Path.GetFileNameWithoutExtension(path) + " (synced)" + Path.GetExtension(path);
        }

        private static string Resolve(XElement config, string directory, string name) {
            return ExcelTools.Tooling.ToolPaths.Resolve(directory, ExcelTools.Tooling.ToolPaths.ExpandTemplate(Required(config, name)));
        }

        private static string Required(XElement config, string name) {
            var value = (string)config.Attribute(name);
            if (String.IsNullOrWhiteSpace(value)) {
                throw new InvalidDataException("Missing " + config.Name + "/@" + name);
            }
            return value;
        }
    }
}