using ExcelDna.Integration;

namespace Sample.AddIn.Standard.Functions
{
    public class CalculateSinus
    {
        // ExcelDna makes calling the Excel API easy:
        // XlCall.Excel works like Excel4, but you just pass the parameters
        // - no need for horrible XLOPERs
        public static double CalcSin(double angle)
        {
            return (double)XlCall.Excel(XlCall.xlfSin, angle);
        }
    }
}