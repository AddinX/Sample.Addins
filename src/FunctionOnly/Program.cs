using System;
using ExcelDna.Integration;
using ExcelDna.Logging;
using NetOffice.ExcelApi;

namespace Sample.AddIn.FunctionOnly
{
    public class Program : IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {
                // Settings the path to the Add-In
                AddinContext.Settings.AddinPath = (string)XlCall.Excel(XlCall.xlGetName);
                
                // The Excel Application object
                AddinContext.ExcelApp = new Application(null, ExcelDnaUtil.Application);
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
