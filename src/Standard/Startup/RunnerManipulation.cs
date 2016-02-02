using AddinX.Bootstrap.Contract;
using Autofac;
using Sample.AddIn.Standard.Manipulation;

namespace Sample.AddIn.Standard.Startup
{
    internal class RunnerManipulation : IRunner
    {
        public void Execute(IRunnerMain bootstrap)
        {
            var bootstrapper = (Bootstrapper)bootstrap;

            bootstrapper?.Builder.RegisterType<TableManipulation>();
        }
    }
}