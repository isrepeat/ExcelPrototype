using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Collections.Generic;

namespace ExcelTools.TableSync.Plugins.Personnel {
    public sealed class PersonnelPlugin : Contracts.ISyncPlugin {
        private static readonly Regex TaxId = new Regex(@"\A[0-9]{10}\z", RegexOptions.CultureInvariant);
        private static readonly Regex NameSuffix = new Regex(@"\s*\([0-9]{10}\)\s*\z", RegexOptions.CultureInvariant);

        public Contracts.SyncPlan BuildPlan(Contracts.SyncContext context) {
            var source = context.ReadTable("shpo", "ОС", null, "ІПН");
            var list = context.ReadTable("movement", "Список", "Список", "ІПН");
            var absences = context.ReadTable("movement", "Відсутні", "ТВ", "ІПН");
            var history = context.ReadTable("registry", "Історія ID", "tbIDHistory", "Попередній ІД");
            var sourceKey = source.Column("ІПН");
            var sourceName = source.Column("Прізвище, ім’я, по батькові");
            var byId = source.Rows.Where(row => TaxId.IsMatch(row.ValueAt(sourceKey)))
                .ToLookup(row => row.ValueAt(sourceKey), StringComparer.Ordinal);
            var byName = source.Rows.ToLookup(row => Name(row.ValueAt(sourceName)), StringComparer.Ordinal);
            var groups = new[] { list, absences }.SelectMany(table => {
                var key = table.Column("ІПН");
                var name = table.Column("ПІБ");
                return table.Rows.Select(row => new { Table = table, Row = row, Key = key, Name = name });
            }).Where(item => TaxId.IsMatch(item.Row.ValueAt(item.Key)))
                .ToLookup(item => item.Row.ValueAt(item.Key), StringComparer.Ordinal);
            var names = groups.ToDictionary(group => group.Key, group => group.Select(item => Name(item.Row.ValueAt(item.Name)))
                .Where(name => name.Length > 0).Distinct(StringComparer.Ordinal).ToArray(), StringComparer.Ordinal);
            var keysByName = names.SelectMany(pair => pair.Value.Select(name => Tuple.Create(name, pair.Key)))
                .ToLookup(pair => pair.Item1, StringComparer.Ordinal);
            var registry = new Registry(history);
            var plan = new Contracts.SyncPlan { SourceBook = source.Book };
            foreach (var group in groups) {
                var key = group.Key;
                var person = Unique(byId[key]);
                var replacement = !byId.Contains(key);
                if (replacement) {
                    var candidates = names[key];
                    if (candidates.Length == 1 && keysByName[candidates[0]].Count() == 1) {
                        person = Unique(byName[candidates[0]]);
                    }
                }
                if (person == null) {
                    if (byId.Contains(key) || names[key].Any(name => byName.Contains(name))) {
                        plan.Groups.Add(new Contracts.SyncGroup { Description = key, Conflict = "Неоднозначное сопоставление с ШПО." });
                    }
                    continue;
                }
                var newKey = person.ValueAt(sourceKey);
                var newName = Name(person.ValueAt(sourceName));
                if (newName.Length == 0 || !TaxId.IsMatch(newKey) || byId[newKey].Count() != 1
                    || (replacement && groups.Contains(newKey))) {
                    plan.Groups.Add(new Contracts.SyncGroup { Description = key, Conflict = "Пустое ФИО, недопустимый или занятый ІПН." });
                    continue;
                }
                var change = new Contracts.SyncGroup { Description = replacement ? key + " → " + newKey : key };
                if (replacement && !registry.Record(change, key, newKey, newName)) {
                    change.Conflict = "Переход конфликтует с историей PeopleRegistry.";
                    plan.Groups.Add(change);
                    continue;
                }
                foreach (var item in group) {
                    if (replacement) {
                        Add(change, item.Table, item.Row.Number, item.Key, newKey);
                    }
                    var before = item.Row.ValueAt(item.Name);
                    if (item.Table == list && !before.StartsWith("=", StringComparison.Ordinal) && Name(before) != newName) {
                        Add(change, list, item.Row.Number, item.Name, person.ValueAt(sourceName));
                    }
                }
                if (change.Edits.Count > 0) {
                    plan.Groups.Add(change);
                }
            }
            return plan;
        }

        private static Contracts.SyncRow Unique(IEnumerable<Contracts.SyncRow> rows) {
            var candidates = rows.Take(2).ToArray();
            return candidates.Length == 1 ? candidates[0] : null;
        }

        private static string Name(string value) {
            return value.StartsWith("=", StringComparison.Ordinal) ? "" : NameSuffix.Replace(value, "").Trim();
        }

        private static void Add(Contracts.SyncGroup group, Contracts.SyncTable table, int row, int column, string value) {
            group.Edits.Add(new Contracts.SyncEdit(table, row, column, value));
        }

        private sealed class Registry {
            private readonly Contracts.SyncTable table;
            private readonly ILookup<string, Contracts.SyncRow> previous;
            private readonly ILookup<string, Contracts.SyncRow> aliases;
            private readonly int oldColumn;
            private readonly int current;
            private readonly int name;
            private readonly int origin;
            private readonly int date;
            private readonly string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            private int next;

            public Registry(Contracts.SyncTable table) {
                this.table = table;
                oldColumn = table.Column("Попередній ІД");
                current = table.Column("ІД");
                name = table.Column("ПІБ");
                origin = table.Column("Звідки");
                date = table.Column("Дата видачі");
                next = table.LastRow;
                previous = table.Rows.ToLookup(row => row.ValueAt(oldColumn), StringComparer.OrdinalIgnoreCase);
                aliases = table.Rows.ToLookup(row => row.ValueAt(current), StringComparer.OrdinalIgnoreCase);
            }

            public bool Record(Contracts.SyncGroup group, string oldKey, string newKey, string fullName) {
                var old = Unique(previous[oldKey]);
                var existing = Unique(previous[newKey]);
                if (old != null && existing != null && old.ValueAt(current) == newKey
                    && existing.ValueAt(current) == newKey && Name(existing.ValueAt(name)) == fullName) {
                    return true;
                }
                if (old == null || old.ValueAt(current) != oldKey || Name(old.ValueAt(name)) != fullName
                    || previous.Contains(newKey) || aliases.Contains(newKey)
                    || aliases[oldKey].Any(row => String.IsNullOrWhiteSpace(row.ValueAt(oldColumn))
                        || previous[row.ValueAt(oldColumn)].Count() != 1)) {
                    return false;
                }
                foreach (var alias in aliases[oldKey]) {
                    Add(group, table, alias.Number, current, newKey);
                }
                next++;
                Add(group, table, next, oldColumn, newKey);
                Add(group, table, next, current, newKey);
                Add(group, table, next, name, fullName);
                Add(group, table, next, origin, old.ValueAt(origin));
                Add(group, table, next, date, timestamp);
                return true;
            }
        }
    }
}