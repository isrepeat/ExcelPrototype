using System;
using System.IO;
using System.Linq;
using System.Text;

using ExcelTools.OpenXml;
using ExcelTools.Tooling;
using ExcelTools.Tooling.Console;

namespace ExcelTools.TableDiff {
    internal static class Program {
        private static int Main(string[] args) {
            return ConsoleTool.Run(args, "TableDiff.log", delegate {
                string configFile = null;
                string profileName = null;
                string oldFile = null;
                string newFile = null;
                var validate = false;
                for (var index = 0; index < args.Length; index++) {
                    var argument = args[index];
                    if (argument == "--help") {
                        Console.WriteLine("TableDiff.exe [--config file.xml] [--profile name | old.xlsx new.xlsx] [--validate] [--no-pause]");
                        Console.WriteLine("Config names are searched beside the executable, in resources, then in ../resources.");
                        return;
                    }
                    else if (argument == "--no-pause") {
                    }
                    else if (argument == "--validate") {
                        validate = true;
                    }
                    else if (argument == "--config") {
                        if (configFile != null || ++index >= args.Length || String.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("--", StringComparison.Ordinal)) {
                            throw new ArgumentException("Specify --config once followed by an XML file.");
                        }
                        configFile = args[index];
                    }
                    else if (argument == "--profile") {
                        if (profileName != null || ++index >= args.Length || String.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("--", StringComparison.Ordinal)) {
                            throw new ArgumentException("Specify --profile once followed by a profile name.");
                        }
                        profileName = args[index];
                    }
                    else if (argument.StartsWith("--", StringComparison.Ordinal)) {
                        throw new ArgumentException("Unknown argument: " + argument);
                    }
                    else if (oldFile == null) {
                        oldFile = argument;
                    }
                    else if (newFile == null) {
                        newFile = argument;
                    }
                    else {
                        throw new ArgumentException("Specify only two workbook paths.");
                    }
                }
                var config = ConfigurationReader.Read(configFile, profileName, oldFile, newFile);
                Console.WriteLine("Configuration: " + config.Path);
                Console.WriteLine("Profile: " + config.Profile);
                Console.WriteLine("Old: " + config.Old);
                Console.WriteLine("New: " + config.New);
                Console.WriteLine("Output: " + config.Output);
                foreach (var comparison in config.Comparisons) {
                    Console.WriteLine(
                        "Sheets: {0} -> {1}; key: {2}; threshold: {3}",
                        comparison.OldSheet,
                        comparison.NewSheet,
                        comparison.Key,
                        comparison.Threshold
                    );
                }
                if (validate) {
                    Console.WriteLine("Configuration and paths validated; workbook contents were not read.");
                    return;
                }
                Console.WriteLine("Reading old workbook...");
                var before = WorkbookReader.Read(config.Old, config.Comparisons.Select(c => c.OldSheet));
                Console.WriteLine("Reading new workbook...");
                var after = WorkbookReader.Read(config.New, config.Comparisons.Select(c => c.NewSheet));
                Console.WriteLine("Comparing...");
                var differences = TableComparer.Compare(before, after, config.Comparisons);
                var document = HtmlDiffWriter.Generate(config, differences);
                Directory.CreateDirectory(Path.GetDirectoryName(config.Output));
                var temporary = OutputPublisher.CreateTemporaryPath(config.Output, ".html");
                try {
                    File.WriteAllText(temporary, document, new UTF8Encoding(false));
                    OutputPublisher.Publish(temporary, config.Output);
                }
                finally {
                    if (File.Exists(temporary)) {
                        File.Delete(temporary);
                    }
                }
                Console.WriteLine(
                    "Changed: {0}; added: {1}; deleted: {2}",
                    differences.Sum(s => s.Count("change")),
                    differences.Sum(s => s.Count("add")),
                    differences.Sum(s => s.Count("delete"))
                );
                Console.WriteLine("Created: " + config.Output);
            });
        }
    }
}