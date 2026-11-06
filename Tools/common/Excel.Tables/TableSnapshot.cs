using System;
using System.Collections.Generic;

namespace ExcelTools.Tables {
    public sealed class RowSnapshot {
        public int Number;
        public string[] Values;

        public string ValueAt(int column) {
            return column < Values.Length ? Values[column] : "";
        }
    }

    public sealed class SheetSnapshot {
        public string Name;
        public int Width = 1;
        public List<RowSnapshot> Rows = new List<RowSnapshot>();
    }

    public sealed class WorkbookSnapshot {
        public List<SheetSnapshot> Sheets = new List<SheetSnapshot>();

        public SheetSnapshot FindSheet(string name) {
            foreach (var sheet in Sheets) {
                if (String.Equals(sheet.Name, name, StringComparison.Ordinal)) {
                    return sheet;
                }
            }
            return null;
        }
    }

    public static class CellAddress {
        public static int RowNumber(string address) {
            return Int32.Parse(address.TrimStart('$').TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()).TrimStart('$'));
        }

        public static string ColumnName(int column) {
            var name = "";
            while (column > 0) {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }

        public static int ColumnNumber(string address) {
            var column = 0;
            foreach (var character in address.TrimStart('$')) {
                if (character < 'A' || character > 'Z') {
                    break;
                }
                column = checked(column * 26 + character - 'A' + 1);
            }
            if (column < 1 || column > 16384) {
                throw new ArgumentException("Invalid Excel column: " + address);
            }
            return column;
        }
    }
}