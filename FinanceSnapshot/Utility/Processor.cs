using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace FinanceSnapshot.Utility
{
    internal class Processor
    {
        Worksheet? sheet;
        Range? excelRange;
        public void ProcessExcel(Workbook workbook)
        {
            int numSheets = workbook.Sheets.Count;

            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Worksheet sheet = workbook.Sheets[sheetNum];
                Range excelRange = sheet.UsedRange;
                object[,] valueArray = excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                // do something (i.e. format, apply formulas, compress, truncate etc.) with values

                DisplayExcel(sheet, excelRange, valueArray);
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey(true);
            }
        }
        public void DisplayExcel(Worksheet sheet, Range excelRange, object[,] valueArray)
        {
            for (int r = 1; r <= valueArray.GetLength(0); r++)
            {
                for (int c = 1; c <= valueArray.GetLength(1); c++)
                {
                    string cellValue = valueArray[r, c]?.ToString() ?? string.Empty;
                    Console.Write($"{cellValue}\t");
                }
            }
        }
    }
}
