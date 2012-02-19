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
                var export = new ResourceExport(options.InputDirectory, options.Language, options.OutputFile, options.Recursive);
                export.Export();
            }
        }
    }
}
