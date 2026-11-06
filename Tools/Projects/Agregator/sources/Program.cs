using System;

namespace ExcelTools.Agregator {
    internal static class Program {
        [STAThread]
        private static int Main(string[] args) {
            return ExcelTools.Tooling.Console.ConsoleTool.Run(args, "Agregator.log", delegate {
                    var validate = false;
                    string output = null;
                    string configFile = null;
                    for (var i = 0; i < args.Length; i++) {
                        var argument = args[i];
                        if (argument == "--help") {
                            Console.WriteLine("Agregator.exe [--config file.xml] [--validate] [--no-pause] [output.xlsx]");
                        Console.WriteLine("Config names are searched beside the executable, in resources, then in ../resources.");
                        Console.WriteLine("Explicit config paths are relative to the executable; source/output paths are relative to the selected config.");
                            Console.WriteLine("Existing output is replaced after success.");
                            return;
                        }
                        if (argument == "--validate") {
                            validate = true;
                        }
                        else if (argument == "--no-pause") {
                        }
                        else if (argument == "--config") {
                            if (configFile != null) {
                                throw new ArgumentException("Specify --config only once.");
                            }
                            if (++i >= args.Length || String.IsNullOrWhiteSpace(args[i]) || args[i].StartsWith("--", StringComparison.Ordinal)) {
                                throw new ArgumentException("Specify an XML file after --config.");
                            }
                            configFile = args[i];
                        }
                        else if (argument.StartsWith("--", StringComparison.Ordinal)) {
                            throw new ArgumentException("Unknown option: " + argument);
                        }
                        else if (output != null) {
                            throw new ArgumentException("Specify only one optional output path.");
                        }
                        else {
                            output = argument;
                        }
                    }
                    var config = ConfigurationReader.Read(configFile, output);
                    Console.WriteLine("Configuration: " + config.Path);
                    Console.WriteLine("Copy mode: " + config.Mode);
                    if (validate) {
                        Console.WriteLine("Paths validated. Input workbooks: " + config.Sources.Count);
                        Console.WriteLine("Output: " + config.Output);
                    }
                    else {
                        AggregationService.Execute(config);
                    }
                });
        }
    }
}