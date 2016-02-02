using System.Threading;
using AddinX.Logging;
using Autofac;
using NetOffice.ExcelApi;
using Sample.AddIn.Standard.Controller;

namespace Sample.AddIn.Standard
{
    internal static class AddinContext
    {
        private static MainController ctrls;

        private static AddInSettings settings;
        
        public static AddInSettings Settings => settings ?? (settings = new AddInSettings());

        public static CancellationTokenSource TokenCancellationSource { get; set; }

        public static IContainer Container { get; set; }

        public static Application ExcelApp { get; set; }

        public static MainController MainController => ctrls ?? (ctrls = Container.Resolve<MainController>());
    }
}