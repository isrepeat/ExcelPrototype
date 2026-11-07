using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace ExcelTools.Tooling {
    public static class OutputPublisher {
        public static void PublishBatch(IDictionary<string, string> files) {
            if (files.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != files.Count) {
                throw new ArgumentException("Output paths must be distinct.");
            }
            var backups = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var published = new List<string>();
            try {
                foreach (var output in files.Values) {
                    if (File.Exists(output)) {
                        var backup = CreateTemporaryPath(output, ".backup");
                        File.Move(output, backup);
                        backups.Add(output, backup);
                    }
                }
                foreach (var file in files) {
                    File.Move(file.Key, file.Value);
                    published.Add(file.Value);
                }
            }
            catch (Exception publishError) {
                var errors = new List<Exception> { publishError };
                foreach (var output in published) {
                    try {
                        File.Delete(output);
                    }
                    catch (Exception error) {
                        errors.Add(error);
                    }
                }
                foreach (var backup in backups) {
                    try {
                        File.Move(backup.Value, backup.Key);
                    }
                    catch (Exception error) {
                        errors.Add(new IOException("Recovery file retained at " + backup.Value, error));
                    }
                }
                throw new AggregateException("Cannot publish output files.", errors);
            }
            foreach (var backup in backups.Values) {
                try {
                    File.Delete(backup);
                }
                catch (Exception error) {
                    Console.Error.WriteLine("Previous copy retained at " + backup + ": " + error.Message);
                }
            }
        }

        public static string CreateTemporaryPath(string output, string extension) {
            return output + "." + Guid.NewGuid().ToString("N") + ".tmp" + extension;
        }

        public static void Publish(string temporary, string output) {
            if (!File.Exists(output)) {
                File.Move(temporary, output);
                return;
            }
            // Rename files within the same directory for SMB compatibility.
            var backup = output + ".previous-" + Guid.NewGuid().ToString("N") + ".tmp";
            File.Move(output, backup);
            try {
                File.Move(temporary, output);
            }
            catch (Exception publishError) {
                try {
                    File.Move(backup, output);
                }
                catch (Exception restoreError) {
                    throw new AggregateException("Cannot restore previous output. Recovery file: " + backup, publishError, restoreError);
                }
                throw;
            }
            try {
                File.Delete(backup);
            }
            catch (Exception error) {
                Console.Error.WriteLine("Warning: output published; previous copy retained at " + backup + ": " + error.Message);
            }
        }

    }
}