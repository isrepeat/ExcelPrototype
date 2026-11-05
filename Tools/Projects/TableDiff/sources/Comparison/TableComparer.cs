using System;
using System.Linq;
using System.Collections.Generic;

using ExcelTools.Tables;

namespace ExcelTools.TableDiff {
    internal sealed class RowChange {
        public string Kind;
        public RowSnapshot Before;
        public RowSnapshot After;
        public bool Ambiguous;
    }

    internal sealed class SheetDiff {
        public string Name;
        public string OldSheet;
        public string NewSheet;
        public string Key;
        public double Threshold;
        public int Width;
        public int OldColumnOffset;
        public int NewColumnOffset;
        public RowSnapshot OldHeader;
        public RowSnapshot NewHeader;
        public int[] OldColumns;
        public int[] NewColumns;
        public int OldCount;
        public int NewCount;
        public List<RowChange> Changes = new List<RowChange>();

        public int Count(string kind) {
            return Changes.Count(c => c.Kind == kind);
        }
    }

    internal static class TableComparer {
        private sealed class TableRange {
            public int Start;
            public int Width;
        }

        private sealed class Candidate {
            public int Index;
            public double Score;
            public int Distance;
        }

        public static List<SheetDiff> Compare(WorkbookSnapshot before, WorkbookSnapshot after, List<SheetComparison> comparisons) {
            var result = new List<SheetDiff>();
            foreach (var comparison in comparisons) {
                var oldSheet = before.FindSheet(comparison.OldSheet);
                var newSheet = after.FindSheet(comparison.NewSheet);
                if (oldSheet == null || newSheet == null) {
                    throw new InvalidOperationException("Selected worksheet not found in " + (oldSheet == null ? "old" : "new")
                        + " workbook: " + (oldSheet == null ? comparison.OldSheet : comparison.NewSheet));
                }
                var difference = CompareSheet(oldSheet, newSheet, comparison.Key, comparison.Threshold);
                difference.OldSheet = comparison.OldSheet;
                difference.NewSheet = comparison.NewSheet;
                difference.Name = comparison.OldSheet == comparison.NewSheet ? comparison.OldSheet
                    : comparison.OldSheet + " → " + comparison.NewSheet;
                difference.Key = comparison.Key;
                difference.Threshold = comparison.Threshold;
                result.Add(difference);
            }
            return result;
        }

        private static SheetDiff CompareSheet(SheetSnapshot before, SheetSnapshot after, string key, double threshold) {
            // Detect each table independently and compare columns relative to its left boundary.
            var oldRange = FindRange(before, key);
            var newRange = FindRange(after, key);
            if (oldRange != null || newRange != null) {
                if ((oldRange == null && before.Rows.Count > 0) || (newRange == null && after.Rows.Count > 0)) {
                    throw new InvalidOperationException("Missing key header '" + key + "' in one version of sheet '" + before.Name + "'.");
                }
                before = SelectRange(before, oldRange ?? new TableRange { Width = newRange.Width });
                after = SelectRange(after, newRange ?? new TableRange { Width = oldRange.Width });
            }
            var result = new SheetDiff {
                Name = before.Name,
                Width = Math.Max(before.Width, after.Width),
                OldColumnOffset = oldRange == null ? 0 : oldRange.Start,
                NewColumnOffset = newRange == null ? 0 : newRange.Start,
                OldHeader = before.Rows.FirstOrDefault(r => r.Values.Contains(key, StringComparer.Ordinal)),
                NewHeader = after.Rows.FirstOrDefault(r => r.Values.Contains(key, StringComparer.Ordinal)),
                OldCount = before.Rows.Count,
                NewCount = after.Rows.Count
            };
            if (result.OldHeader != null && result.NewHeader != null) {
                AlignColumns(ref before, ref after, result);
            }

            // Remove exact matches first, including reordered and repeated rows.
            var pool = new Dictionary<string, Queue<int>>(StringComparer.Ordinal);
            for (var index = 0; index < after.Rows.Count; index++) {
                var signature = Signature(after.Rows[index], result.Width);
                Queue<int> indices;
                if (!pool.TryGetValue(signature, out indices)) {
                    indices = new Queue<int>();
                    pool.Add(signature, indices);
                }
                indices.Enqueue(index);
            }
            var usedBefore = new HashSet<int>();
            var usedAfter = new HashSet<int>();
            for (var index = 0; index < before.Rows.Count; index++) {
                Queue<int> indices;
                if (pool.TryGetValue(Signature(before.Rows[index], result.Width), out indices) && indices.Count > 0) {
                    usedBefore.Add(index);
                    usedAfter.Add(indices.Dequeue());
                }
            }

            // Pair remaining table rows by key, similarity, proximity and descending candidate index.
            if (result.OldHeader != null && result.NewHeader != null) {
                var keyColumn = Array.IndexOf(result.OldHeader.Values, key);
                var candidates = new Dictionary<string, List<int>>(StringComparer.Ordinal);
                for (var index = 0; index < after.Rows.Count; index++) {
                    var row = after.Rows[index];
                    if (usedAfter.Contains(index) || row.Number <= result.NewHeader.Number) {
                        continue;
                    }
                    List<int> indices;
                    var identity = row.ValueAt(keyColumn);
                    if (!candidates.TryGetValue(identity, out indices)) {
                        indices = new List<int>();
                        candidates.Add(identity, indices);
                    }
                    indices.Add(index);
                }
                for (var index = 0; index < before.Rows.Count; index++) {
                    var row = before.Rows[index];
                    var identity = row.ValueAt(keyColumn);
                    List<int> indices;
                    if (usedBefore.Contains(index) || row.Number <= result.OldHeader.Number || identity.Length == 0
                        || !candidates.TryGetValue(identity, out indices)) {
                        continue;
                    }
                    var ranked = new List<Candidate>();
                    foreach (var targetIndex in indices) {
                        if (usedAfter.Contains(targetIndex)) {
                            continue;
                        }
                        var target = after.Rows[targetIndex];
                        var active = 0;
                        var equal = 0;
                        for (var column = 0; column < result.Width; column++) {
                            if (result.OldColumns != null && (result.OldColumns[column] < 0 || result.NewColumns[column] < 0)) {
                                continue;
                            }
                            var x = row.ValueAt(column);
                            var y = target.ValueAt(column);
                            if (x.Length > 0 || y.Length > 0) {
                                active++;
                                if (x == y) {
                                    equal++;
                                }
                            }
                        }
                        ranked.Add(new Candidate {
                            Index = targetIndex,
                            Score = (double)equal / Math.Max(1, active),
                            Distance = Math.Abs(row.Number - target.Number)
                        });
                    }
                    ranked = ranked.OrderByDescending(c => c.Score).ThenBy(c => c.Distance).ThenByDescending(c => c.Index).ToList();
                    if (ranked.Count > 0 && ranked[0].Score >= threshold) {
                        var targetIndex = ranked[0].Index;
                        result.Changes.Add(new RowChange {
                            Kind = "change",
                            Before = row,
                            After = after.Rows[targetIndex],
                            Ambiguous = ranked.Count > 1 && ranked[0].Score == ranked[1].Score
                        });
                        usedBefore.Add(index);
                        usedAfter.Add(targetIndex);
                    }
                }
            }

            // Compare unmatched rows outside the recognized table by their original row number.
            var remaining = new Dictionary<int, int>();
            for (var index = 0; index < after.Rows.Count; index++) {
                if (!usedAfter.Contains(index)) {
                    remaining.Add(after.Rows[index].Number, index);
                }
            }
            for (var index = 0; index < before.Rows.Count; index++) {
                var row = before.Rows[index];
                if (usedBefore.Contains(index) || (result.OldHeader != null && row.Number > result.OldHeader.Number)) {
                    continue;
                }
                int targetIndex;
                if (remaining.TryGetValue(row.Number, out targetIndex)) {
                    result.Changes.Add(new RowChange { Kind = "change", Before = row, After = after.Rows[targetIndex] });
                    usedBefore.Add(index);
                    usedAfter.Add(targetIndex);
                }
            }
            for (var index = 0; index < before.Rows.Count; index++) {
                if (!usedBefore.Contains(index)) {
                    result.Changes.Add(new RowChange { Kind = "delete", Before = before.Rows[index] });
                }
            }
            for (var index = 0; index < after.Rows.Count; index++) {
                if (!usedAfter.Contains(index)) {
                    result.Changes.Add(new RowChange { Kind = "add", After = after.Rows[index] });
                }
            }
            if (result.OldCount + result.Count("add") - result.Count("delete") != result.NewCount) {
                throw new InvalidOperationException("Row count balance failed: " + result.Name);
            }
            result.Changes = result.Changes.OrderBy(c => (c.After ?? c.Before).Number).ToList();
            return result;
        }

        private static TableRange FindRange(SheetSnapshot sheetSnapshot, string key) {
            var header = sheetSnapshot.Rows.FirstOrDefault(r => r.Values.Contains(key, StringComparer.Ordinal));
            if (header == null) {
                return null;
            }
            var start = Array.IndexOf(header.Values, key);
            var end = start;
            while (start > 0 && header.ValueAt(start - 1).Length > 0) {
                start--;
            }
            while (end + 1 < header.Values.Length && header.ValueAt(end + 1).Length > 0) {
                end++;
            }
            return new TableRange { Start = start, Width = end - start + 1 };
        }

        private static void AlignColumns(ref SheetSnapshot before, ref SheetSnapshot after, SheetDiff difference) {
            var oldNames = difference.OldHeader.Values;
            var newNames = difference.NewHeader.Values;
            if (oldNames.SequenceEqual(newNames)) {
                return;
            }
            if (oldNames.Distinct(StringComparer.Ordinal).Count() != oldNames.Length
                || newNames.Distinct(StringComparer.Ordinal).Count() != newNames.Length) {
                throw new InvalidOperationException("Cannot map changed columns with duplicate headers in sheet '" + difference.Name + "'.");
            }
            var names = newNames.Concat(oldNames.Where(name => !newNames.Contains(name, StringComparer.Ordinal))).ToArray();
            var oldMap = names.Select(name => Array.IndexOf(oldNames, name)).ToArray();
            var newMap = names.Select(name => Array.IndexOf(newNames, name)).ToArray();
            difference.OldColumns = oldMap.Select(index => index < 0 ? -1 : difference.OldColumnOffset + index).ToArray();
            difference.NewColumns = newMap.Select(index => index < 0 ? -1 : difference.NewColumnOffset + index).ToArray();
            before = ProjectColumns(before, oldMap);
            after = ProjectColumns(after, newMap);
            difference.Width = names.Length;
            difference.OldHeader = before.Rows.First(row => row.Number == difference.OldHeader.Number);
            difference.NewHeader = after.Rows.First(row => row.Number == difference.NewHeader.Number);
        }

        private static SheetSnapshot ProjectColumns(SheetSnapshot source, int[] columns) {
            var result = new SheetSnapshot { Name = source.Name, Width = columns.Length };
            foreach (var row in source.Rows) {
                result.Rows.Add(new RowSnapshot {
                    Number = row.Number,
                    Values = columns.Select(column => column < 0 ? "" : row.ValueAt(column)).ToArray()
                });
            }
            return result;
        }

        private static SheetSnapshot SelectRange(SheetSnapshot sheetSnapshot, TableRange tableRange) {
            var result = new SheetSnapshot { Name = sheetSnapshot.Name, Width = tableRange.Width };
            foreach (var row in sheetSnapshot.Rows) {
                var values = Enumerable.Range(tableRange.Start, tableRange.Width).Select(row.ValueAt).ToArray();
                if (values.Any(value => value.Length > 0)) {
                    result.Rows.Add(new RowSnapshot { Number = row.Number, Values = values });
                }
            }
            return result;
        }

        private static string Signature(RowSnapshot rowSnapshot, int width) {
            // Length prefixes avoid collisions when cell text contains separator characters.
            return String.Concat(Enumerable.Range(0, width).Select(c => {
                var value = rowSnapshot.ValueAt(c);
                return value.Length + ":" + value;
            }));
        }
    }
}