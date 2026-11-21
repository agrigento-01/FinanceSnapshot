using FinanceSnapshot.Utility;
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinanceSnapshot
{
    public class Program
    {
        public void OpenExcelFile(string filePath)
        {
            try
            {
                Excel.Application _excelApp = new Excel.Application();
                Excel.Workbook workbook = _excelApp.Workbooks.Open(filePath);

                // Perform operations on the workbook
                Processor processor = new Processor();
                processor.ProcessExcel(workbook);

                workbook.Close(false, filePath, null);
                _excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                // todo
            }
        }
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            Program program = new Program();

            program.OpenExcelFile(@"C:\Users\John\Desktop\Finance Sheet 2026.xlsx");
        }
    }
}