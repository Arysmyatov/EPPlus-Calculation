using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace APPLust_excel_calculation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Open file
            var filePath = "EPPlus-calculation.xlsx";
            var fileInfo = new FileInfo(filePath);

            // Open Xls file
            using (var package = new ExcelPackage(fileInfo))
            {
                var currencyRateSheet = package.Workbook.Worksheets["currency rate"];
                currencyRateSheet.Cells[1, 1].Value = 29;

                currencyRateSheet = package.Workbook.Worksheets[1];

                Console.WriteLine(currencyRateSheet.Cells[1,2].Value);

                package.Save();
            }

            // Open Xls file
            using (var package = new ExcelPackage(fileInfo))
            {
                var currencyRateSheet = package.Workbook.Worksheets[1];
                currencyRateSheet.Calculate();
                Console.WriteLine(currencyRateSheet.Cells["B1"].Value);
            }            

            Console.ReadKey();            
        }
    }
}
