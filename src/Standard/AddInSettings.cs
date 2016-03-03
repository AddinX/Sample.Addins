using AddinX.Configuration.Implementation;

namespace Sample.AddIn.Standard
{
    internal class AddInSettings
    {
        public string AddinName =>  AddinContext.ConfigManager.AppSettings["addin:name"];

        public string AddinPath { get; set; }
    }
}