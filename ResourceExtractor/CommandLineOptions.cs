using System.Text;
using CommandLine;
using CommandLine.Text;

namespace ResourceExtractor
{
    class CommandLineOptions
    {
        [Option("e", "", HelpText = "Export resources to xls (default)")]
        public bool Export = false;

        [Option("i", "", HelpText = "Import xls to resources")]
        public bool Import = false;

        [Option("d", "", Required = true, HelpText = "Input resource directory / xls to read.")]
        public string InputDirectory = null;

        [Option("l", "", HelpText = "The language of the main resource.")]
        public string Language = null;

        [Option("x", "", Required = true, HelpText = "The import / export xls file")]
        public string XlsFile = null;

        [HelpOption(HelpText = "Display this help screen.")]
        public string GetUsage()
        {
            var help = new HelpText("ResourceExtractor");
            help.Copyright = new CopyrightInfo("Luuk Sommers", 2012);
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Contact me at my blog: http://sameproblemmorecode.blogspot.com");
            help.AddPreOptionsLine("Check for updates at: https://github.com/luuksommers");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Examples:");
            help.AddPreOptionsLine("ResourceExtractor.exe -e -d C:\\Project\\Resources -x C:\\Project\\ResourceExtract.xls -l en");
            help.AddPreOptionsLine("ResourceExtractor.exe -i -d C:\\Project\\Resources -x C:\\Project\\ResourceExtract.xls -l en");
            help.AddPreOptionsLine("");
            help.AddOptions(this);
            return help;
        }
    }
}