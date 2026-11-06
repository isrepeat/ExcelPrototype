using System;
using System.Linq;
using System.Collections.Generic;

namespace ExcelTools.Tables {
    public sealed class TableMatch {
        public RowSnapshot Target { get; private set; }
        public RowSnapshot Source { get; private set; }
        public bool ByFallback { get; private set; }

        internal TableMatch(RowSnapshot target, RowSnapshot source, bool byFallback) {
            Target = target;
            Source = source;
            ByFallback = byFallback;
        }
    }

    public static class TableMatching {
        public static RowSnapshot FindHeader(SheetSnapshot sheet, IEnumerable<string> names, int? rowNumber) {
            var headers = sheet.Rows.Where(row => (!rowNumber.HasValue || row.Number == rowNumber.Value)
                && names.All(name => row.Values.Contains(name, StringComparer.Ordinal))).Take(2).ToArray();
            if (headers.Length != 1) {
                throw new InvalidOperationException("Expected one header row with the configured columns on sheet: " + sheet.Name);
            }
            return headers[0];
        }

        public static int FindColumn(RowSnapshot header, string name) {
            var indices = Enumerable.Range(0, header.Values.Length)
                .Where(index => String.Equals(header.Values[index], name, StringComparison.Ordinal)).ToArray();
            if (indices.Length != 1) {
                throw new InvalidOperationException("Expected one column named '" + name + "' in header row " + header.Number);
            }
            return indices[0];
        }

        public static List<RowSnapshot> DataRows(SheetSnapshot sheet, RowSnapshot header, int? lastRow) {
            if (lastRow.HasValue && lastRow.Value <= header.Number) {
                throw new InvalidOperationException("Last data row must follow the header: " + sheet.Name);
            }
            return sheet.Rows.Where(row => row.Number > header.Number && (!lastRow.HasValue || row.Number <= lastRow.Value)).ToList();
        }

        public static Dictionary<int, TableMatch> Match(
            List<RowSnapshot> target,
            List<RowSnapshot> source,
            int targetKey,
            int sourceKey,
            int? targetFallback,
            int? sourceFallback,
            bool matchAllTargetRows = false) {
            var targetKeys = Index(target, targetKey);
            var sourceKeys = Index(source, sourceKey);
            var targetNames = targetFallback.HasValue ? Index(target, targetFallback.Value) : null;
            var sourceNames = sourceFallback.HasValue ? Index(source, sourceFallback.Value) : null;
            var result = new Dictionary<int, TableMatch>();
            foreach (var row in target) {
                var key = row.ValueAt(targetKey);
                if (!IsMatchValue(key)) {
                    continue;
                }
                var uniqueTarget = targetKeys[key].Count == 1;
                List<RowSnapshot> candidates;
                if (sourceKeys.TryGetValue(key, out candidates)) {
                    if (candidates.Count == 1 && (uniqueTarget || matchAllTargetRows)) {
                        result.Add(row.Number, new TableMatch(row, candidates[0], false));
                    }
                    continue;
                }

                if (!uniqueTarget || targetNames == null || sourceNames == null) {
                    continue;
                }
                var name = row.ValueAt(targetFallback.Value);
                if (!IsMatchValue(name) || targetNames[name].Count != 1
                    || !sourceNames.TryGetValue(name, out candidates) || candidates.Count != 1) {
                    continue;
                }
                var sourceRow = candidates[0];
                var newKey = sourceRow.ValueAt(sourceKey);
                if (IsMatchValue(newKey) && sourceKeys[newKey].Count == 1 && !targetKeys.ContainsKey(newKey)) {
                    result.Add(row.Number, new TableMatch(row, sourceRow, true));
                }
            }
            return result;
        }

        private static Dictionary<string, List<RowSnapshot>> Index(IEnumerable<RowSnapshot> rows, int column) {
            return rows.Where(row => IsMatchValue(row.ValueAt(column)))
                .GroupBy(row => row.ValueAt(column), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        }

        private static bool IsMatchValue(string value) {
            return !String.IsNullOrWhiteSpace(value) && !value.StartsWith("=", StringComparison.Ordinal);
        }
    }
}