using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace FinanceSnapshot
{
    internal class ExcelReaderInterop
    {
        Application _excelApp;

        public ExcelReaderInterop()
        {
            _excelApp = new Application();
        }
    }
}