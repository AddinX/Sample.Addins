using AddinX.Logging;
using Sample.AddIn.Standard.Manipulation;

namespace Sample.AddIn.Standard.Controller
{
    public class ReportController
    {
        private readonly TableManipulation tableManipulation;
        private readonly ILogger logger;

        public ReportController(TableManipulation tableManipulation, ILogger logger)
        {
            this.tableManipulation = tableManipulation;
            this.logger = logger;
        }

        public void CreateTable()
        {
            logger.Warn("Inside create table method");
            tableManipulation.WriteRange();
        }
    }
}