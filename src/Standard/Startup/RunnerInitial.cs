using AddinX.Bootstrap.Contract;
using AddinX.Logging;
using Autofac;
using ILogger = AddinX.Logging.ILogger;

namespace Sample.AddIn.Standard.Startup
{
    internal class RunnerInitial : IRunner
    {
        public void Execute(IRunnerMain bootstrap)
        {
            var bootstrapper = (Bootstrapper)bootstrap;

            // Excel Application
            bootstrapper?.Builder.RegisterInstance(AddinContext.ExcelApp).ExternallyOwned();

            // Ribbon
            bootstrapper?.Builder.RegisterInstance(new AddinRibbon());

            // ILogger
            bootstrapper?.Builder.RegisterInstance<ILogger>(new SerilogLogger());

            // Security

        }
    }
}