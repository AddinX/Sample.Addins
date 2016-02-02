using AddinX.Logging;
using ExcelDna.Logging;

namespace Sample.AddIn.Standard.Controller
{
    public class SampleController
    {
        private readonly ILogger logger;

        public SampleController(ILogger logger)
        {
            this.logger = logger;
        }

        public void ShowMessage()
        {
            logger.Debug("Inside show message method");
            LogDisplay.Clear();
            LogDisplay.WriteLine("Button Clicked");
            LogDisplay.Show();
        }
    }
}