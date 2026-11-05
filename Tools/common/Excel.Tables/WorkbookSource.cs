using System.Collections.Generic;

namespace ExcelTools.Tables {
    public sealed class WorkbookSource {
        public string Path;
        public List<string> Sheets = new List<string>();
    }
}