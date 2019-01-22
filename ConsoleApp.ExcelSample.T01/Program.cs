using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp.ExcelSample.T01
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = Path.GetTempPath() + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "-csharp-Excel.xls";
            string logFile = Path.GetTempPath() + "ConsoleApp.ExcelSample.T01.log";
            string message = string.Empty;

            try
            {
                LogToFile(logFile, "Try to create Excel instance");
                Excel.Application xlApp = new Excel.Application();
                LogToFile(logFile, "Done to create Excel instance");

                if (xlApp == null)
                {
                    message = "Excel is not properly installed!!";
                    Console.WriteLine(message);
                    LogToFile(logFile, message);

                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "ID";
                xlWorkSheet.Cells[1, 2] = "Name";
                xlWorkSheet.Cells[1, 3] = "Address";
                xlWorkSheet.Cells[2, 1] = "abcdefghigklmn";
                xlWorkSheet.Cells[2, 2] = "One";
                xlWorkSheet.Cells[2, 3] = "Three";
                xlWorkSheet.Cells[3, 1] = "2";
                xlWorkSheet.Cells[3, 2] = "abcdefghigklmn";
                xlWorkSheet.Cells[3, 3] = "cccccc";

                xlWorkSheet.Columns.AutoFit();

                LogToFile(logFile, "Try to save Excel file");
                xlWorkBook.SaveAs(excelFilePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                LogToFile(logFile, "Done to save Excel file");

                LogToFile(logFile, "Try to close Excel object");
                xlWorkBook.Close(true, misValue, misValue);
                LogToFile(logFile, "Done to close Excel object");

                LogToFile(logFile, "Try to quit Excel application");
                xlApp.Quit();
                LogToFile(logFile, "Done to quit Excel application");

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                message = "Excel file created , you can find the file " + excelFilePath;
                Console.WriteLine(message);
                LogToFile(logFile, message);
            }
            catch (Exception ex)
            {
                LogToFile(logFile, "Encounter exception, message is " + ex.ToString());
            }
        }

        private static void LogToFile(string logFile, string message)
        {
            using (StreamWriter file = File.AppendText(logFile))
            {
                file.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:sss") + "   " + message);
            }
        }

    }
}
