namespace Sample.AddIn.Standard.Controller
{
    public class MainController
    {
        public MainController(ReportController report, SampleController sample)
        {
            Report = report;
            Sample = sample;
        }

        public SampleController Sample { get; private set; }

        public ReportController Report { get; private set; }
    }
}