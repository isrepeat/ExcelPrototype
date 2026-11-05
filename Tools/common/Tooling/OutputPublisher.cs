using System;
using System.IO;

namespace ExcelTools.Tooling {
    public static class OutputPublisher {
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