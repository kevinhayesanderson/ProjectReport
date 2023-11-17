using Services;
using Utilities;

namespace Actions
{
    [ActionName("InOutEntry")]
    internal class InOutEntryAction(bool run, string inputFolder, string time, ILogger logger, ReadService readService, ExportService exportService) : IAction
    {
        public string InputFolder => inputFolder;

        public bool Run => run;

        public bool Execute()
        {
            bool res = true;
            return res;
        }
    }
}