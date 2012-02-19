using System.Text;
using CommandLine;
using CommandLine.Text;

namespace ResourceExtractor
{
    class CommandLineOptions
    {
        [Option("i", "input", Required = true, HelpText = "Input directory to read.")]
        public string InputDirectory = null;

        [Option("l", "language", HelpText = "The language of the main resource.")]
        public string Language = null;

        [Option("o", "output", HelpText = "The output file.")]
        public string OutputFile = null;

        [Option("r", "recursive", HelpText = "Search recursively through the input directory")]
        public bool Recursive = false;

        [HelpOption(HelpText = "Display this help screen.")]
        public string GetUsage()
        {
            HelpText help = new HelpText("ResourceExtractor");
            help.Copyright = new CopyrightInfo("Luuk Sommers", 2012);
            help.AddPreOptionsLine("Contact me at my blog: http://sameproblemmorecode.blogspot.com");
            help.AddPreOptionsLine("Check for updates at: https://github.com/luuksommers");
            help.AddOptions(this);
            return help;
        }
    }
}