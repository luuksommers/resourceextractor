using System;
using CommandLine;

namespace ResourceExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new CommandLineOptions();
            ICommandLineParser parser = new CommandLineParser();
            if (parser.ParseArguments(args, options, Console.Out))
            {
                if (options.Export)
                {
                    var export = new ResourceExport(options.InputDirectory, options.Language, options.OutputFile, options.Recursive);
                    export.Export();
                }
                else
                {
                    var importer = new ResourceImport(options.InputDirectory, options.Language, options.OutputFile);
                    importer.Import();
                }
            }
        }
    }
}
