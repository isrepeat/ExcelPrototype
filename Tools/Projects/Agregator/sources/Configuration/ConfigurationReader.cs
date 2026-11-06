using System;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;

namespace ExcelTools.Agregator {
    internal sealed class Configuration {
        public string Path;
        public string Output;
        public string Mode;
        public List<ExcelTools.Tables.WorkbookSource> Sources = new List<ExcelTools.Tables.WorkbookSource>();
    }

    internal static class ConfigurationReader {
        private static string Required(XElement element, string name) {
            var value = (string)element.Attribute(name);
            if (String.IsNullOrWhiteSpace(value)) {
                throw new InvalidDataException("Missing configuration attribute: " + element.Name + "/@" + name);
            }
            return value;
        }

        public static Configuration Read(string configFile, string outputOverride) {
            var directory = AppDomain.CurrentDomain.BaseDirectory;
            var config = new Configuration();
            config.Path = ExcelTools.Tooling.ToolPaths.FindConfiguration(directory, configFile ?? "Agregator.xml");
            directory = System.IO.Path.GetDirectoryName(config.Path);
            var document = XDocument.Load(config.Path);
            var root = document.Root;
            if (root == null || root.Name != "aggregator") {
                throw new InvalidDataException("Expected aggregator configuration root.");
            }
            config.Output = ExcelTools.Tooling.ToolPaths.Resolve(directory, outputOverride ?? Required(root, "output"));
            config.Mode = Required(root, "copyMode");
            if (config.Mode != "values" && config.Mode != "sheet" && config.Mode != "xmlStyles") {
                throw new InvalidDataException("Invalid copyMode: " + config.Mode + ". Expected values, sheet or xmlStyles.");
            }
            var sources = root.Element("sources");
            if (sources == null || !sources.Elements("source").Any()) {
                throw new InvalidDataException("No source directories configured.");
            }
            foreach (var source in sources.Elements("source")) {
                var sourceDirectory = ExcelTools.Tooling.ToolPaths.Resolve(directory, Required(source, "dir"));
                if (!source.Elements("workbook").Any()) {
                    throw new InvalidDataException("No source workbooks configured: " + sourceDirectory);
                }
                foreach (var workbook in source.Elements("workbook")) {
                    var item = new ExcelTools.Tables.WorkbookSource();
                    item.Path = ExcelTools.Tooling.ToolPaths.Resolve(sourceDirectory, Required(workbook, "file"));
                    item.Sheets.AddRange(workbook.Elements("sheet").Select(e => Required(e, "name")));
                    if (item.Sheets.Count == 0) {
                        throw new InvalidDataException("No worksheets selected: " + item.Path);
                    }
                    if (!File.Exists(item.Path)) {
                        throw new FileNotFoundException("Input workbook does not exist: " + item.Path);
                    }
                    if (String.Equals(item.Path, config.Output, StringComparison.OrdinalIgnoreCase)) {
                        throw new InvalidDataException("Input and output paths must differ: " + item.Path);
                    }
                    if (config.Mode == "xmlStyles" && !String.Equals(System.IO.Path.GetExtension(item.Path), ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                        throw new InvalidDataException("xmlStyles requires .xlsx sources: " + item.Path);
                    }
                    config.Sources.Add(item);
                }
            }
            if (!String.Equals(System.IO.Path.GetExtension(config.Output), ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Output must have the .xlsx extension: " + config.Output);
            }
            if (Directory.Exists(config.Output)) {
                throw new IOException("Output path is a directory: " + config.Output);
            }
            if (!Directory.Exists(System.IO.Path.GetDirectoryName(config.Output))) {
                throw new DirectoryNotFoundException("Output directory does not exist: " + config.Output);
            }
            return config;
        }

    }
}