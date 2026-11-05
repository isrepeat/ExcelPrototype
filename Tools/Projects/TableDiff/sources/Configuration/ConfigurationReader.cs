using System;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;

using ExcelTools.Tooling;

namespace ExcelTools.TableDiff {
    internal sealed class Configuration {
        public string Path;
        public string Profile;
        public string Old;
        public string New;
        public string Output;
        public List<SheetComparison> Comparisons = new List<SheetComparison>();
    }

    internal sealed class SheetComparison {
        public string OldSheet;
        public string NewSheet;
        public string Key;
        public double Threshold;
    }

    internal static class ConfigurationReader {
        private static string Required(XElement element, string name) {
            if (element == null || String.IsNullOrWhiteSpace((string)element.Attribute(name))) {
                throw new InvalidDataException("Missing configuration attribute: " + (element == null ? "source" : element.Name.LocalName) + "/@" + name);
            }
            return (string)element.Attribute(name);
        }

        public static Configuration Read(string configFile, string profileName, string oldFile, string newFile) {
            var path = ToolPaths.FindConfiguration(AppDomain.CurrentDomain.BaseDirectory, configFile ?? "TableDiff.xml");
            var root = XDocument.Load(path).Root;
            if (root == null || root.Name != "tableDiff") {
                throw new InvalidDataException("Expected tableDiff configuration root.");
            }
            var container = root.Element("profiles");
            if (container == null || !container.Elements("profile").Any()) {
                throw new InvalidDataException("Specify at least one profiles/profile element; each profile must contain old, new and comparisons.");
            }
            if (root.Element("old") != null || root.Element("new") != null || root.Element("comparisons") != null
                || root.Attribute("key") != null || root.Attribute("threshold") != null || root.Attribute("output") != null) {
                throw new InvalidDataException("Move file paths, output and sheet comparisons into profiles/profile.");
            }

            // Parse all profile rules, but access workbook files only for the selected profile.
            var profiles = new List<Configuration>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var filePairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var element in container.Elements("profile")) {
                var profile = ReadProfile(path, element);
                if (!names.Add(profile.Profile)) {
                    throw new InvalidDataException("Duplicate profile name: " + profile.Profile);
                }
                var identity = profile.Old.Length + ":" + profile.Old + profile.New.Length + ":" + profile.New;
                if (!filePairs.Add(identity)) {
                    throw new InvalidDataException("Multiple profiles use the same ordered file pair: " + profile.Old + " -> " + profile.New);
                }
                profiles.Add(profile);
            }

            Configuration selected;
            if (profileName != null) {
                if (oldFile != null || newFile != null) {
                    throw new ArgumentException("Choose either --profile or two workbook paths.");
                }
                selected = profiles.FirstOrDefault(p => String.Equals(p.Profile, profileName, StringComparison.OrdinalIgnoreCase));
                if (selected == null) {
                    throw new InvalidDataException("Unknown profile: " + profileName + ". Available profiles: " + String.Join(", ", profiles.Select(p => p.Profile)));
                }
            }
            else if (oldFile != null || newFile != null) {
                if (oldFile == null || newFile == null) {
                    throw new ArgumentException("Specify both old and new workbook paths.");
                }
                var oldPath = ToolPaths.Resolve(Environment.CurrentDirectory, oldFile);
                var newPath = ToolPaths.Resolve(Environment.CurrentDirectory, newFile);
                selected = profiles.FirstOrDefault(p =>
                    String.Equals(p.Old, oldPath, StringComparison.OrdinalIgnoreCase)
                    && String.Equals(p.New, newPath, StringComparison.OrdinalIgnoreCase));
                if (selected == null) {
                    throw new InvalidDataException("No profile matches the ordered file pair: " + oldPath + " -> " + newPath);
                }
            }
            else if (profiles.Count == 1) {
                selected = profiles[0];
            }
            else {
                throw new ArgumentException("Multiple profiles configured. Specify --profile or two workbook paths. Available profiles: "
                    + String.Join(", ", profiles.Select(p => p.Profile)));
            }
            ValidatePaths(selected);
            return selected;
        }

        private static Configuration ReadProfile(string path, XElement element) {
            var directory = System.IO.Path.GetDirectoryName(path);
            var config = new Configuration {
                Path = path,
                Profile = Required(element, "name"),
                Old = ToolPaths.Resolve(directory, Required(element.Element("old"), "file")),
                New = ToolPaths.Resolve(directory, Required(element.Element("new"), "file")),
                Output = ToolPaths.Resolve(directory, Required(element, "output"))
            };
            if (element.Attribute("key") != null || element.Attribute("threshold") != null) {
                throw new InvalidDataException("Specify key and threshold on each comparison in profile: " + config.Profile);
            }
            var comparisons = element.Element("comparisons");
            if (comparisons == null) {
                throw new InvalidDataException("Missing comparisons in profile: " + config.Profile);
            }
            var pairs = new HashSet<string>(StringComparer.Ordinal);
            foreach (var comparisonElement in comparisons.Elements("comparison")) {
                var comparison = new SheetComparison {
                    OldSheet = Required(comparisonElement.Element("old"), "sheet"),
                    NewSheet = Required(comparisonElement.Element("new"), "sheet"),
                    Key = Required(comparisonElement, "key")
                };
                if (!Double.TryParse(Required(comparisonElement, "threshold"), NumberStyles.Float, CultureInfo.InvariantCulture, out comparison.Threshold)
                    || Double.IsNaN(comparison.Threshold) || comparison.Threshold <= 0 || comparison.Threshold > 1) {
                    throw new InvalidDataException("Invalid threshold in profile '" + config.Profile + "' for '" + comparison.OldSheet + "' -> '"
                        + comparison.NewSheet + "': expected a number greater than 0 and not greater than 1.");
                }
                var identity = comparison.OldSheet.Length + ":" + comparison.OldSheet + comparison.NewSheet.Length + ":" + comparison.NewSheet;
                if (!pairs.Add(identity)) {
                    throw new InvalidDataException("Duplicate sheet pair in profile '" + config.Profile + "': '" + comparison.OldSheet
                        + "' -> '" + comparison.NewSheet + "'.");
                }
                config.Comparisons.Add(comparison);
            }
            if (config.Comparisons.Count == 0) {
                throw new InvalidDataException("Specify at least one comparison in profile: " + config.Profile);
            }
            return config;
        }

        private static void ValidatePaths(Configuration configuration) {
            foreach (var input in new[] { configuration.Old, configuration.New }) {
                if (!File.Exists(input)) {
                    throw new FileNotFoundException("Input workbook does not exist: " + input, input);
                }
                var extension = System.IO.Path.GetExtension(input);
                if (!String.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                    && !String.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidDataException("Expected an unencrypted .xlsx or .xlsm workbook: " + input);
                }
                if (String.Equals(input, configuration.Output, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidDataException("The report cannot overwrite an input workbook.");
                }
            }
            if (!String.Equals(System.IO.Path.GetExtension(configuration.Output), ".html", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Output must have the .html extension.");
            }
            if (Directory.Exists(configuration.Output)) {
                throw new IOException("Output path is a directory: " + configuration.Output);
            }
        }
    }
}