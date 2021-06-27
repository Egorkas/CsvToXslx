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
        public static void CsvConvert(string fullInputPathFile, string outputPathDirectory)
        {
            string workSheetsName = "Bank";
            var firstRowIsHeader = false;
            var format = new ExcelTextFormat();
            format.Delimiter = ';';
            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
            format.EOL = "\n";
            format.Encoding = Encoding.GetEncoding(1251);
            //format.Encoding = Encoding.GetEncoding(fullInputPathFile);

            var totalRowCounter = File.ReadLines(fullInputPathFile).Count();
            var outputName = Path.GetFileNameWithoutExtension(fullInputPathFile) + ".xlsx";

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(outputPathDirectory + "\\" + outputName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(workSheetsName);

                worksheet.Cells["A1"].LoadFromText((new FileInfo(fullInputPathFile)), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                package.Save();
            }
        }

    }
}
