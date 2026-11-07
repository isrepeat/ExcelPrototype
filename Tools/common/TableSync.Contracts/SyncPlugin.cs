using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelTools.TableSync.Contracts {
    public interface ISyncPlugin {
        SyncPlan BuildPlan(SyncContext context);
    }

    public sealed class SyncContext {
        private readonly Func<string, string, string, string, SyncTable> read;

        public SyncContext(Func<string, string, string, string, SyncTable> read) {
            this.read = read;
        }

        public SyncTable ReadTable(string book, string sheet, string table, string keyColumn) {
            return read(book, sheet, table, keyColumn);
        }
    }

    public sealed class SyncTable {
        public string Book { get; private set; }
        public string Sheet { get; private set; }
        public string Table { get; private set; }
        public int LastRow { get; private set; }
        public ReadOnlyCollection<string> Headers { get; private set; }
        public ReadOnlyCollection<SyncRow> Rows { get; private set; }

        public SyncTable(string book, string sheet, string table, int lastRow, IEnumerable<string> headers, IEnumerable<SyncRow> rows) {
            Book = book;
            Sheet = sheet;
            Table = table;
            LastRow = lastRow;
            Headers = Array.AsReadOnly(headers.ToArray());
            Rows = Array.AsReadOnly(rows.ToArray());
        }

        public int Column(string name) {
            return Enumerable.Range(0, Headers.Count).Single(index => Headers[index] == name);
        }
    }

    public sealed class SyncRow {
        public int Number { get; private set; }
        private readonly string[] values;

        public SyncRow(int number, IEnumerable<string> values) {
            Number = number;
            this.values = values.ToArray();
        }

        public string ValueAt(int column) {
            return column < values.Length ? values[column] : "";
        }
    }

    public sealed class SyncEdit {
        public SyncTable SourceTable { get; set; }
        public SyncRow SourceRow { get; set; }
        public int SourceColumn { get; set; }
        public string Reason { get; set; }
        public SyncTable Table { get; private set; }
        public int Row { get; private set; }
        public int Column { get; private set; }
        public string Value { get; private set; }

        public SyncEdit(SyncTable table, int row, int column, string value) {
            Table = table;
            Row = row;
            Column = column;
            Value = value;
        }
    }

    public sealed class SyncGroup {
        public string Description { get; set; }
        public string Conflict { get; set; }
        public bool DerivedConflict { get; set; }
        public List<SyncTable> ConflictTables { get; private set; }
        public List<SyncEdit> Edits { get; private set; }

        public SyncGroup() {
            Edits = new List<SyncEdit>();
            ConflictTables = new List<SyncTable>();
        }
    }

    public sealed class SyncPlan {
        public string SourceBook { get; set; }
        public List<SyncGroup> Groups { get; private set; }

        public SyncPlan() {
            Groups = new List<SyncGroup>();
        }
    }
}