using NetOffice.ExcelApi;

namespace Sample.AddIn.FunctionOnly
{
    internal static class AddinContext
    {
        private static AddInSettings settings;

        public static AddInSettings Settings => settings ?? (settings = new AddInSettings());
       
        public static Application ExcelApp { get; set; }
    }
}