using System;
using System.Text.RegularExpressions;
using System.Globalization;

using ExcelTools.Tables;

namespace ExcelTools.OpenXml {
    internal static class FormulaTranslator {
        // Skip quoted strings, sheet names and structured references before translating A1 references.
        private static readonly Regex References = new Regex(
            @"""(?:[^""]|"""")*""|'(?:[^']|'')*'|\[[^\]]*\]|(?<![\w.])(?<column>\$?[A-Z]{1,3})(?<row>\$?[1-9][0-9]*)(?![\w.(!])|(?<![\w.])(?<columns>\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![\w.])|(?<![\w.])(?<rows>\$?[1-9][0-9]*:\$?[1-9][0-9]*)(?![\w.])"
        );

        public static string Translate(string formula, string origin, string destination) {
            var columnOffset = CellAddress.ColumnNumber(destination) - CellAddress.ColumnNumber(origin);
            var rowOffset = Row(destination) - Row(origin);
            return References.Replace(formula, match => {
                if (match.Groups["column"].Success) {
                    return ShiftColumn(match.Groups["column"].Value, columnOffset) + ShiftRow(match.Groups["row"].Value, rowOffset);
                }
                if (match.Groups["columns"].Success) {
                    var columns = match.Value.Split(':');
                    return ShiftColumn(columns[0], columnOffset) + ":" + ShiftColumn(columns[1], columnOffset);
                }
                if (match.Groups["rows"].Success) {
                    var rows = match.Value.Split(':');
                    return ShiftRow(rows[0], rowOffset) + ":" + ShiftRow(rows[1], rowOffset);
                }
                return match.Value;
            });
        }

        private static int Row(string address) {
            return Int32.Parse(Regex.Match(address, @"[0-9]+$").Value, CultureInfo.InvariantCulture);
        }

        private static string ShiftColumn(string column, int offset) {
            if (column.StartsWith("$", StringComparison.Ordinal)) {
                return column;
            }
            var number = CellAddress.ColumnNumber(column) + offset;
            if (number < 1 || number > 16384) {
                throw new InvalidOperationException("Shared formula column is outside the worksheet.");
            }
            return CellAddress.ColumnName(number);
        }

        private static string ShiftRow(string row, int offset) {
            if (row.StartsWith("$", StringComparison.Ordinal)) {
                return row;
            }
            var number = Int32.Parse(row, CultureInfo.InvariantCulture) + offset;
            if (number < 1 || number > 1048576) {
                throw new InvalidOperationException("Shared formula row is outside the worksheet.");
            }
            return number.ToString(CultureInfo.InvariantCulture);
        }
    }
}