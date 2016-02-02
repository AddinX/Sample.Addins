using ExcelDna.Integration;

namespace Sample.AddIn.FunctionOnly.Functions
{
    public class DescribeArguments
    {

        // This function returns a string that describes its argument.
        // For arguments defined as object type, this shows all the possible types that may be received.
        // Also try this function after changing the 
        // [ExcelArgument(AllowReference=true)] attribute.
        // In that case we allow references to be passed (registerd as type R). 
        // By default the function will be registered not
        // to receive references AllowReference=false (type P).
        [ExcelFunction(Description = "Describes the value passed to the function.", IsMacroType = true)]
        public static string Describe([ExcelArgument(AllowReference = false)] object arg)
        {
            if (arg is double)
                return "Double: " + (double) arg;
            else if (arg is string)
                return "String: " + (string) arg;
            else if (arg is bool)
                return "Boolean: " + (bool) arg;
            else if (arg is ExcelError)
                return "ExcelError: " + arg.ToString();
            else if (arg is object[,])
                // The object array returned here may contain a mixture of different types,
                // reflecting the different cell contents.
                return string.Format("Array[{0},{1}]", ((object[,]) arg).GetLength(0), ((object[,]) arg).GetLength(1));
            else if (arg is ExcelMissing)
                return "Missing";
            else if (arg is ExcelEmpty)
                return "Empty";
            else if (arg is ExcelReference)
                return "Reference: " + XlCall.Excel(XlCall.xlfReftext, arg, true);
            else
                return "!?Unheard Of";
        }
    }
}
