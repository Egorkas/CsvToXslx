using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CsvToXslx
{
    class CsvConverterToXlsx
    {
        public static void CsvConvert(string file)
        {
            string csvFileName = @"d:\PZ\CsvToXslx\CsvToXslx\bin\Debug\net5.0\CsvHelper\1.csv";
            string xlsxFileName = @"d:\PZ\CsvToXslx\CsvToXslx\bin\Debug\net5.0\CsvHelper\CsvToXlsx\1.xlsx";
            string workSheetsName = "Bank";
            var firstRowIsHeader = false;
            var format = new ExcelTextFormat();
            format.Delimiter = ';';
            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
            format.EOL = "\n";
            format.Encoding = Encoding.GetEncoding(1251);
            var totalRowCounter = File.ReadLines(csvFileName).Count();
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(xlsxFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(workSheetsName);

                worksheet.Cells["A1"].LoadFromText((new FileInfo(csvFileName)), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                package.Save();
            }
        }

    }
}
