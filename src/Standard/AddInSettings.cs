using AddinX.Configuration.Implementation;

namespace Sample.AddIn.Standard
{
    internal class AddInSettings
    {
        public string AddinName => new ConfigurationManagerWrapper().AppSettings["addin:name"];

        public string AddinPath { get; set; }
    }
}