using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;

namespace ExcelTools.Tooling {
    public static class ToolPaths {
        public static string ExpandTemplate(string template, string targetName = null) {
            var date = DateTime.Now.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture);
            return Regex.Replace(template, @"\{(?:date|targetName)\}", match => {
                if (String.Equals(match.Value, "{date}", StringComparison.OrdinalIgnoreCase)) {
                    return date;
                }
                return targetName ?? match.Value;
            }, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        public static string FindConfiguration(string directory, string fileName) {
            if (Path.IsPathRooted(fileName) || !String.Equals(Path.GetFileName(fileName), fileName, StringComparison.Ordinal)) {
                var explicitPath = Resolve(directory, fileName);
                if (!File.Exists(explicitPath)) {
                    throw new FileNotFoundException("Configuration file does not exist: " + explicitPath, explicitPath);
                }
                return explicitPath;
            }
            var candidates = new[] {
                Resolve(directory, fileName),
                Resolve(directory, Path.Combine("resources", fileName)),
                Resolve(directory, Path.Combine("..", "resources", fileName))
            };
            foreach (var candidate in candidates) {
                if (File.Exists(candidate)) {
                    return candidate;
                }
            }
            throw new FileNotFoundException("Configuration file not found. Searched: " + String.Join("; ", candidates), fileName);
        }

        public static string Resolve(string directory, string path) {
            return Path.GetFullPath(Path.Combine(directory, path));
        }
    }
}