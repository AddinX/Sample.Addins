using System.Runtime.InteropServices;
using System.Threading;
using AddinX.Ribbon.Contract;
using AddinX.Ribbon.Contract.Command;
using AddinX.Ribbon.ExcelDna;

namespace Sample.AddIn.Standard
{
    [ComVisible(true)]
    public class AddinRibbon : RibbonFluent
    {
        protected override void CreateFluentRibbon(IRibbonBuilder build)
        {
            build.CustomUi.Ribbon.Tabs(tab =>
                tab.AddTab("MY TAB").SetId("CustomTab")
                    .Groups(g =>
                    {
                        g.AddGroup("Forms").SetId("SampleGroup")
                            .Items(i =>
                            {
                                i.AddButton("Test").SetId("TestCmd")
                                    .LargeSize().ImageMso("HappyFace").ShowLabel()
                                    .Screentip("First button for test purpose");
                                i.AddButton("Table").SetId("TableCmd")
                                    .NormalSize();
                            });
                    }));

        }

        protected override void CreateRibbonCommand(IRibbonCommands cmds)
        {
            cmds.AddButtonCommand("TestCmd")
                .IsEnabled(() => true).IsVisible(() => true)
                .Action(() => AddinContext.MainController.Sample.ShowMessage());
            
            cmds.AddButtonCommand("TableCmd").IsVisible(() => true)
                .Action(() => AddinContext.MainController.Report.CreateTable());
        }

        public override void OnClosing()
        {
            AddinContext.TokenCancellationSource.Cancel();
            Thread.Sleep(100);

            AddinContext.ExcelApp.DisposeChildInstances(true);
            AddinContext.ExcelApp = null;

            AddinContext.Container.Dispose();
            AddinContext.Container = null;
        }

        public override void OnOpening()
        {
            // Register to events
            AddinContext.ExcelApp.SheetActivateEvent += (e) => RefreshRibbon();
            AddinContext.ExcelApp.SheetChangeEvent += (a,e) => RefreshRibbon();
        }

        private void RefreshRibbon()
        {
            Ribbon?.Invalidate();
        }
    }
}