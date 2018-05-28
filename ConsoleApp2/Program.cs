using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ConsoleApp2
{
    internal static class Program
    {
        static void Main()
        {
            //int rCnt;
            //int cCnt;

            var res = ExcelImport(@"d:\sample.xls");

        }

        private static DataTable ExcelImport(string filePath)
        {
            var xlApp = new Application();

            var xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];

            var range = xlWorkSheet.UsedRange;
            var rw = range.Rows.Count;

            var dt = new DataTable();
            dt.Columns.Add("A");
            dt.Columns.Add("B");

            for (var rCnt = 1; rCnt <= rw; rCnt++)
            {
                var newCustomersRow = dt.NewRow();
                newCustomersRow["A"] = (string)(range.Cells[rCnt, 1] as Range)?.Value2.ToString(); 
                newCustomersRow["B"] = (string)(range.Cells[rCnt, 2] as Range)?.Value2.ToString();

                dt.Rows.Add(newCustomersRow);
            }

            xlWorkBook.Close(true);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return dt;
        }
    }
}
