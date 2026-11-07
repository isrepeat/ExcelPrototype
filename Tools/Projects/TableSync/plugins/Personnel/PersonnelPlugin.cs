using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Collections.Generic;

namespace ExcelTools.TableSync.Plugins.Personnel {
    public sealed class PersonnelPlugin : Contracts.ISyncPlugin {
        private static readonly Regex TaxId = new Regex(@"\A[0-9]{10}\z", RegexOptions.CultureInvariant);
        private static readonly Regex NameSuffix = new Regex(@"\s*\((?<id>[0-9]{10})\)\s*\z", RegexOptions.CultureInvariant);

        public Contracts.SyncPlan BuildPlan(Contracts.SyncContext context) {
            var source = context.ReadTable("shpo", "ОС", null, "ІПН");
            var list = context.ReadTable("movement", "Список", "Список", "ІПН");
            var absences = context.ReadTable("movement", "Відсутні", "ТВ", "ІПН");
            var history = context.ReadTable("registry", "Історія ID", "tbIDHistory", "Попередній ІД");
            var sourceKey = source.Column("ІПН");
            var sourceName = source.Column("Прізвище, ім’я, по батькові");
            var byId = source.Rows
                .Where(row => TaxId.IsMatch(row.ValueAt(sourceKey)))
                .ToLookup(
                    row => row.ValueAt(sourceKey),
                    StringComparer.Ordinal
                );

            var entries = new[] { list, absences }
                .SelectMany(table => table.Rows.Select(row => new PersonRow(table, row, table.Column("ІПН"), table.Column("ПІБ"))))
                .Where(item => TaxId.IsMatch(item.Id))
                .ToArray();
            var groups = entries.ToLookup(item => item.Id, StringComparer.Ordinal);
            var sourceEntries = source.Rows.Select(row => new PersonRow(source, row, sourceKey, sourceName)).ToArray();
            var sourceById = sourceEntries.ToLookup(item => item.Id, StringComparer.Ordinal);
            var sourceByName = sourceEntries.ToLookup(item => item.Name, StringComparer.Ordinal);
            var movementByName = entries.ToLookup(item => item.Name, StringComparer.Ordinal);
            var registry = new Registry(history);
            var plan = new Contracts.SyncPlan {
                SourceBook = source.Book
            };

            foreach (var group in groups) {
                var key = group.Key;
                var names = group.Select(item => item.Name)
                    .Where(name => name.Length > 0)
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();
                var relatedSource = sourceById[key].Concat(names.SelectMany(name => sourceByName[name])).Distinct().ToArray();
                var relatedMovement = group.Concat(names.SelectMany(name => movementByName[name])).Distinct().ToArray();
                var issue = group.FirstOrDefault(item => item.Qualifier.Length > 0 && item.Qualifier != item.Id);
                if (issue != null) {
                    Conflict(plan, key, "ІПН в скобках ФИО расходится с колонкой ІПН.", relatedMovement, relatedSource);
                    continue;
                }

                var person = Unique(byId[key]);
                var replacement = !byId.Contains(key);
                if (replacement) {
                    if (names.Length != 1) {
                        Conflict(plan, key, "Один ІПН связан с разными ФИО в РУХ.", relatedMovement, relatedSource);
                        continue;
                    }
                    var candidates = sourceByName[names[0]].ToArray();
                    var wrongQualifier = candidates.Any(item => item.Qualifier.Length > 0 && item.Qualifier != item.Id);
                    if (wrongQualifier) {
                        Conflict(plan, key, "В ШПО ІПН в скобках ФИО расходится с колонкой ІПН.", relatedMovement, candidates);
                        continue;
                    }
                    var explicitlyDifferent = group.Any(item => item.Qualifier.Length > 0)
                        && candidates.Length > 0
                        && candidates.All(item => item.Qualifier.Length > 0 && item.Qualifier != key);
                    if (explicitlyDifferent) {
                        Conflict(plan, key, "Разные люди с одинаковым ФИО: в РУХ и ШПО указаны разные ІПН в скобках.", relatedMovement, candidates);
                        continue;
                    }
                    if (group.Any(item => item.Qualifier.Length > 0)) {
                        if (candidates.Length > 0) {
                            Conflict(plan, key, "РУХ явно уточняет личность через ІПН в скобках. Смена номера по одному ФИО требует ручной проверки.", relatedMovement, candidates);
                        }
                        continue;
                    }
                    if (candidates.Length != 1) {
                        if (candidates.Length > 0) {
                            Conflict(plan, key, "Несколько строк ШПО имеют одинаковое ФИО.", relatedMovement, candidates);
                        }
                        continue;
                    }
                    if (relatedMovement.Any(item => item.Id != key)) {
                        Conflict(plan, key, "Одинаковое ФИО связано с разными ІПН в РУХ. Без уточняющих скобок нельзя определить, ошибочен ли ІПН в РУХ или ШПО.", relatedMovement, candidates);
                        continue;
                    }
                    person = candidates[0].Row;
                }
                if (person == null) {
                    Conflict(plan, key, "В ШПО несколько строк с одним ІПН.", relatedMovement, relatedSource);
                    continue;
                }
                var newKey = person.ValueAt(sourceKey);
                var sourcePerson = new PersonRow(source, person, sourceKey, sourceName);
                if (sourcePerson.Qualifier.Length > 0 && sourcePerson.Qualifier != newKey) {
                    Conflict(plan, key, "В ШПО ІПН в скобках ФИО расходится с колонкой ІПН.", relatedMovement, relatedSource);
                    continue;
                }
                var newName = sourcePerson.Name;
                if (newName.Length == 0
                    || !TaxId.IsMatch(newKey)
                    || byId[newKey].Count() != 1
                    || (replacement && groups.Contains(newKey))) {
                    Conflict(plan, key, "Пустое ФИО, недопустимый, повторяющийся или уже занятый ІПН.", relatedMovement, relatedSource);
                    continue;
                }
                var change = new Contracts.SyncGroup {
                    Description = newName + " · " + (replacement ? key + " → " + newKey : key)
                };
                if (replacement && !registry.Record(change, key, newKey, newName)) {
                    var conflict = Conflict(plan, change.Description, "Переход конфликтует с историей PeopleRegistry.", relatedMovement, relatedSource);
                    conflict.ConflictTables.Clear();
                    conflict.ConflictTables.Add(history);
                    conflict.DerivedConflict = true;
                    conflict.Conflict += Environment.NewLine + registry.Describe(key, newKey);
                    continue;
                }
                foreach (var item in group) {
                    if (replacement) {
                        Add(change, item.Table, item.Row.Number, item.Key, newKey);
                    }
                    var before = item.Row.ValueAt(item.NameColumn);
                    if (!before.StartsWith("=", StringComparison.Ordinal)
                        && item.Name != newName) {
                        var suffix = item.Qualifier.Length > 0 && sourcePerson.Qualifier.Length == 0 ? " (" + item.Qualifier + ")" : "";
                        Add(change, item.Table, item.Row.Number, item.NameColumn, person.ValueAt(sourceName) + suffix);
                    }
                }
                if (change.Edits.Count > 0) {
                    foreach (var edit in change.Edits) {
                        edit.SourceTable = source;
                        edit.SourceRow = person;
                        if (edit.Table == history) {
                            edit.SourceColumn = -1;
                        }
                        else {
                            edit.SourceColumn = edit.Column == edit.Table.Column("ІПН") ? sourceKey : sourceName;
                        }
                        edit.Reason = replacement ? "Замена ІПН " + key + " → " + newKey : "ФИО отличается от ШПО";
                    }
                    plan.Groups.Add(change);
                }
            }
            return plan;
        }

        private sealed class PersonRow {
            public readonly Contracts.SyncTable Table;
            public readonly Contracts.SyncRow Row;
            public readonly int Key;
            public readonly int NameColumn;
            public readonly string Id;
            public readonly string Name;
            public readonly string Qualifier;

            public PersonRow(Contracts.SyncTable table, Contracts.SyncRow row, int key, int name) {
                Table = table;
                Row = row;
                Key = key;
                NameColumn = name;
                Id = row.ValueAt(key);
                var value = row.ValueAt(name);
                Name = PersonnelPlugin.Name(value);
                var match = value.StartsWith("=", StringComparison.Ordinal) ? Match.Empty : NameSuffix.Match(value);
                Qualifier = match.Success ? match.Groups["id"].Value : "";
            }

            public string Describe() {
                return Table.Book + " · " + Table.Sheet + "!" + Row.Number
                    + " | ІПН " + Id + " | " + Row.ValueAt(NameColumn);
            }
        }

        private static Contracts.SyncGroup Conflict(
            Contracts.SyncPlan plan, string key, string reason,
            IEnumerable<PersonRow> movement, IEnumerable<PersonRow> source) {
            var details = movement.Select(item => item.Describe()).Concat(source.Select(item => item.Describe()));
            var conflict = new Contracts.SyncGroup {
                Description = key,
                Conflict = reason + Environment.NewLine + String.Join(Environment.NewLine, details)
            };
            conflict.ConflictTables.AddRange(movement.Select(item => item.Table).Distinct());
            plan.Groups.Add(conflict);
            return conflict;
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
                if (old != null
                    && existing != null
                    && old.ValueAt(current) == newKey
                    && existing.ValueAt(current) == newKey
                    && Name(existing.ValueAt(name)) == fullName) {
                    return true;
                }
                if (old == null
                    || old.ValueAt(current) != oldKey
                    || Name(old.ValueAt(name)) != fullName
                    || previous.Contains(newKey)
                    || aliases.Contains(newKey)
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

            public string Describe(string oldKey, string newKey) {
                var rows = previous[oldKey].Concat(previous[newKey]).Concat(aliases[oldKey]).Concat(aliases[newKey]).Distinct();
                var details = rows.Select(row => table.Book + " · " + table.Sheet + "!" + row.Number
                    + " | " + row.ValueAt(oldColumn) + " → " + row.ValueAt(current) + " | " + row.ValueAt(name)).ToArray();
                return details.Length == 0
                    ? "В истории нет записи для перехода " + oldKey + " → " + newKey
                    : String.Join(Environment.NewLine, details);
            }
        }
    }
}