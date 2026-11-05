using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace ExcelTools.Tooling.Console {
    public static class ConsoleTool {
        public static int Run(string[] args, string logFileName, Action action) {
            global::System.Console.OutputEncoding = new UTF8Encoding(false);
            var watch = Stopwatch.StartNew();
            var success = false;
            try {
                action();
                success = true;
                return 0;
            }
            catch (Exception error) {
                try {
                    var log = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logFileName);
                    File.AppendAllText(log, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine + error + Environment.NewLine, new UTF8Encoding(false));
                    global::System.Console.Error.WriteLine("Error details saved to: " + log);
                }
                catch (Exception logError) {
                    global::System.Console.Error.WriteLine("Cannot write error log: " + logError.Message);
                }
                var aggregate = error as AggregateException;
                if (aggregate != null) {
                    global::System.Console.Error.WriteLine("Error: " + aggregate.Message);
                    foreach (var cause in aggregate.Flatten().InnerExceptions) {
                        global::System.Console.Error.WriteLine("Error: " + cause.Message);
                    }
                }
                else {
                    global::System.Console.Error.WriteLine("Error: " + error.Message);
                }
                return 1;
            }
            finally {
                global::System.Console.WriteLine("Status: " + (success ? "SUCCESS" : "FAILED"));
                global::System.Console.WriteLine("Total elapsed: {0:F3} s", watch.Elapsed.TotalSeconds);
                if (!args.Contains("--no-pause") && !global::System.Console.IsInputRedirected) {
                    global::System.Console.WriteLine("Press any key to close...");
                    global::System.Console.ReadKey(true);
                }
            }
        }
    }
}