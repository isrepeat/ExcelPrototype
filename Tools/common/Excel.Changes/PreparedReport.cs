using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;

namespace ExcelTools.Changes {
    public sealed class RowChange {
        public string Kind;
        public ExcelTools.Tables.RowSnapshot Before;
        public ExcelTools.Tables.RowSnapshot After;
        public bool Ambiguous;
        public string ReviewNote;
    }

    public sealed class SheetDiff {
        public string Book;
        public string Name;
        public string OldSheet;
        public string NewSheet;
        public string Key;
        public double Threshold;
        public int Width;
        public int OldColumnOffset;
        public int NewColumnOffset;
        public ExcelTools.Tables.RowSnapshot OldHeader;
        public ExcelTools.Tables.RowSnapshot NewHeader;
        public int[] OldColumns;
        public int[] NewColumns;
        public int OldCount;
        public int NewCount;
        public List<RowChange> Changes = new List<RowChange>();

        public int Count(string kind) {
            return Changes.Count(c => c.Kind == kind);
        }
    }

    public sealed class ReportWorkbook {
        public List<string> Sheets;
        public string Name;
        public string Before;
        public string After;
    }

    public sealed class SyncPreviewItem {
        public string SourceValue;
        public string Group;
        public string Field;
        public string Before;
        public string After;
        public string Reason;
        public bool Added;
        public List<string> Targets = new List<string>();
        public List<string> Sources = new List<string>();
    }

    public sealed class SyncPreviewSection {
        public string Book;
        public string Sheet;
        public string File;
        public string SourceFile;
        public bool Derived;
        public List<SyncPreviewItem> Items = new List<SyncPreviewItem>();
        public List<string> Conflicts = new List<string>();
    }

    [XmlRoot("preparedReport")]
    public sealed class PreparedReport {
        [XmlAttribute("version")]
        public int Version;
        public string ConfigurationPath;
        public string SyncProfile;
        public string ReportProfile;
        public string Output;
        public string Method;
        public string StagedWorkbook;
        public string OldFile;
        public string NewFile;
        public string SourceFile;
        public bool Preview;
        public List<string> Conflicts;
        public List<SyncPreviewSection> PreviewSections;
        public List<ReportWorkbook> Workbooks;
        public List<SheetDiff> Differences;

        public void Run(string executable) {
            var path = Path.Combine(Path.GetTempPath(), "ExcelChanges-" + System.Guid.NewGuid().ToString("N") + ".xml");
            try {
                using (var stream = File.Open(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    Write(stream, this);
                }
                using (var process = Process.Start(new ProcessStartInfo(executable, "--prepared-report \"" + path + "\" --no-pause") {
                    UseShellExecute = false,
                    WorkingDirectory = Path.GetDirectoryName(executable)
                })) {
                    process.WaitForExit();
                    if (process.ExitCode != 0 || !File.Exists(Output)) {
                        throw new System.InvalidOperationException("TableDiff failed to create report: " + Output);
                    }
                }
            }
            finally {
                File.Delete(path);
            }
        }

        public static void Write(Stream stream, PreparedReport report) {
            report.Version = 1;
            new XmlSerializer(typeof(PreparedReport)).Serialize(stream, report);
        }

        public static PreparedReport Read(string path) {
            using (var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) {
                var report = (PreparedReport)new XmlSerializer(typeof(PreparedReport)).Deserialize(reader);
                if (report.Version != 1) {
                    throw new InvalidDataException("Unsupported prepared report version: " + report.Version);
                }
                return report;
            }
        }
    }
}