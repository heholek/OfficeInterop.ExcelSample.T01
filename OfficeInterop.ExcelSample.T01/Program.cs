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
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string excelFilePath = Path.Combine(Path.GetTempPath(), timeStamp + "-csharp-Excel.xlsx");
            string logFile = Path.Combine(Path.GetTempPath(), timeStamp + "-log.txt");

            LogToFile(logFile, "File path: " + logFile);

            CreateExcelFile(excelFilePath, logFile);

            LogToFile(logFile, "Finished");
        }

        private static void CreateExcelFile(string excelFilePath, string logFile)
        {
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = null;
                Excel.Worksheet xlWorkSheet = null;

                xlWorkBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                xlWorkSheet = xlWorkBook.Worksheets[1];

                for (int i = 1; i < 10; i++)
                {
                    for (int j = 1; j < 11; j++)
                    {
                        xlWorkSheet.Cells[i, j] =  i + "---" + j;
                    }
                }

                xlWorkBook.SaveCopyAs(excelFilePath);

                xlWorkBook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
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
                file.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:sss") + "    " + message);
            }
        }

    }
}
