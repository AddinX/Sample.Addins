using System;
using NetOffice.ExcelApi;

namespace Sample.AddIn.Standard.Manipulation
{
    public class TableManipulation
    {
        private readonly Application excelApp;

        public TableManipulation(Application excelApp)
        {
            this.excelApp = excelApp;
        }

        public void WriteRange()
        {
            // Create the array.
            var myArray = new object[10, 10];

            // Initialize the array.
            for (var i = 0; i < myArray.GetLength(0); i++)
            {
                for (var j = 0; j < myArray.GetLength(1); j++)
                {
                    myArray[i, j] = i + j;
                }
            }

            // Create a Range of the correct size:
            var rows = myArray.GetLength(0);
            var columns = myArray.GetLength(1);
            
            var range = ((Worksheet)excelApp.ActiveSheet).get_Range("A1", Type.Missing);
            range = range.get_Resize(rows, columns);

            // Assign the Array to the Range in one shot:
            range.set_Value(Type.Missing, myArray);
        }
    }
}