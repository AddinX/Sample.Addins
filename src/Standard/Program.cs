using System;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Logging;
using NetOffice.ExcelApi;
using Sample.AddIn.Standard.Startup;

namespace Sample.AddIn.Standard
{
    public class Program : IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {   
                // Settings the path to the Add-In
                AddinContext.Settings.AddinPath = (string) XlCall.Excel(XlCall.xlGetName);
                
                // Token cancellation is useful to close all existing Tasks<> before leaving the application
                AddinContext.TokenCancellationSource = new CancellationTokenSource();

                // The Excel Application object
                AddinContext.ExcelApp = new Application(null, ExcelDnaUtil.Application);
                
                // Start the bootstrapper now
                new Bootstrapper(AddinContext.TokenCancellationSource.Token).Start();
            }
            catch (Exception e)
            {
                LogDisplay.RecordLine(e.Message);
                LogDisplay.RecordLine(e.StackTrace);
                LogDisplay.Show();
            }          
        }

        public void AutoClose()
        {
            throw new NotImplementedException();
        }
    }
}