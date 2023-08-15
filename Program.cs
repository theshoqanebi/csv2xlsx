using System;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace csv2xlsx
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: csv2xlsx.exe <csv file input> <xlsx file output>");
                Console.Write("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            string srcCsvFile = args[0];
            string tgtXlsFile = args[1];

            using (StreamReader reader = new StreamReader(srcCsvFile))
            {
                string line;
                int row = 1;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] values = line.Split(',');
                    int cell = 1;
                    foreach (string item in values)
                    {
                        if (ContainsNumbersOnly(item) || !ContainsNumbers(item))
                        {
                            worksheet.Cells[row, cell] = item;
                        }
                        else
                        {
                            worksheet.Cells[row, cell] = "\u200F" + item;
                        }
                        
                        if (row == 1)
                        {
                            worksheet.Cells[row, cell].AutoFilter();
                            worksheet.Cells[row, cell].Interior.ColorIndex = 31;
                            worksheet.Cells[row, cell].Font.ColorIndex = 2;
                        }
                        worksheet.Cells[row, cell].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        worksheet.Cells[row, cell].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        cell++;
                    }
                    row++;
                }
            }

            workbook.SaveAs(tgtXlsFile);
            workbook.Close();
            excelApp.Quit();
        }
        static bool ContainsNumbersOnly(string input)
        {
            // Regular expression pattern to match only numbers
            string pattern = @"^\d+$";

            return Regex.IsMatch(input, pattern);
        }

        static bool ContainsNumbers(string input)
        {
            return Regex.IsMatch(input, @"\d");
        }
    }
}
