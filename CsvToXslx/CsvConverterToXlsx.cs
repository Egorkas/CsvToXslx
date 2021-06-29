using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CsvToXslx
{
    //public class DataItem
    //{
    //    [Description("Протокол")]
    //    public string Protocol { get; set; }

    //    [Description("IP A")]
    //    public string IpSource { get; set; }

    //    [Description("Port A")]
    //    public string PortSource { get; set; }

    //    [Description("IP B")]
    //    public string IpDestination { get; set; }

    //    [Description("Port B")]
    //    public string PortDestination { get; set; }

    //    [Description("Time")]
    //    public DateTime Time { get; set; }

    //    [Description("Size")]
    //    public int Size { get; set; }

    //    [Description("Сервер")]
    //    public string Server { get; set; }

    //    public string Login { get; set; }

    //    [Description("Time Start")]
    //    public string TimeStart { get; set; }

    //    [Description("Time Stop")]
    //    public string TimeStop { get; set; }

    //    [Description("Провайдер")]
    //    public string Provider { get; set; }

    //    [Description("Логин")]
    //    public string LoginRus { get; set; }

    //    [Description("Подключён")]
    //    public string FirstConnection { get; set; }

    //    [Description("IP")]
    //    public string Ip { get; set; }

    //    [Description("Фамилия")]
    //    public string SecondName { get; set; }

    //    [Description("Имя")]
    //    public string Name { get; set; }

    //    [Description("Отчество")]
    //    public string LastName { get; set; }

    //    [Description("НП / УНП")]
    //    public string Number { get; set; }

    //    [Description("Организация")]
    //    public string Organization { get; set; }

    //    [Description("НП")]
    //    public string City { get; set; }

    //    [Description("ул.")]
    //    public string Street { get; set; }

    //    [Description("д.")]
    //    public string House { get; set; }

    //    [Description("корп.")]
    //    public string Corpus { get; set; }

    //    [Description("кв.")]
    //    public string Flat { get; set; }

    //    [Description("Тел. номер")]
    //    public string PhoneNumber { get; set; }

    //    [Description("E-Mail")]
    //    public string Email { get; set; }

    //    [Description("Доп.")]
    //    public string Additional { get; set; }

    //    [Description("CntBil")]
    //    public string CntBil { get; set; }

    //    [Description("CntOWN")]
    //    public string CntOwn { get; set; }
    //}
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
            var minCellSize = 3;
            var maxCellSize = 27;
            var totalRowCounter = File.ReadLines(fullInputPathFile).Count();
            var outputName = Path.GetFileNameWithoutExtension(fullInputPathFile) + ".xlsx";

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(outputPathDirectory + "\\" + outputName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(workSheetsName);
                worksheet.Cells[1,1].Value = "Протокол"; 
                worksheet.Cells[1,2].Value = "IP A"; 
                worksheet.Cells[1,3].Value = "Port A"; 
                worksheet.Cells[1,4].Value = "IP B"; 
                worksheet.Cells[1,5].Value = "Port B"; 
                worksheet.Cells[1,6].Value = "Time";
                worksheet.Cells[1,7].Value = "Size"; 
                worksheet.Cells[1,8].Value = "Сервер"; 
                worksheet.Cells[1,9].Value = "Сервер"; 
                worksheet.Cells[1,10].Value = "Login"; 
                worksheet.Cells[1,11].Value = "Time Start"; 
                worksheet.Cells[1,12].Value = "Time Stop"; 
                worksheet.Cells[1,13].Value = "Провайдер"; 
                worksheet.Cells[1,14].Value = "Логин"; 
                worksheet.Cells[1,15].Value = "Подключён"; 
                worksheet.Cells[1,16].Value = "IP"; 
                worksheet.Cells[1,17].Value = "Фамилия"; 
                worksheet.Cells[1,18].Value = "Имя"; 
                worksheet.Cells[1,19].Value = "Отчество"; 
                worksheet.Cells[1,20].Value = "НП / УНП"; 
                worksheet.Cells[1,21].Value = "Орг."; 
                worksheet.Cells[1,22].Value = "НП"; 
                worksheet.Cells[1,23].Value = "ул."; 
                worksheet.Cells[1,24].Value = "д."; 
                worksheet.Cells[1,25].Value = "корп."; 
                worksheet.Cells[1,26].Value = "кв."; 
                worksheet.Cells[1,27].Value = "Тел. ном."; 
                worksheet.Cells[1,28].Value = "E-Mail"; 
                worksheet.Cells[1,29].Value = "Доп."; 
                worksheet.Cells[1,30].Value = "CntBil"; 
                worksheet.Cells[1,31].Value = "CntOwn";

                worksheet.Cells["A1:AE1"].Style.Font.Size = 14;
                worksheet.Cells["A1:AE1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1:AE1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#66FFFF"));

                worksheet.Cells["A2"].LoadFromText((new FileInfo(fullInputPathFile)), format, OfficeOpenXml.Table.TableStyles.None, firstRowIsHeader);

                worksheet.Column(6).Style.Numberformat.Format = "dd.MM.yyyy HH:mm:ss";
                worksheet.Column(11).Style.Numberformat.Format = "dd.MM.yyyy HH:mm:ss";
                worksheet.Column(12).Style.Numberformat.Format = "dd.MM.yyyy HH:mm:ss";
                worksheet.Column(15).Style.Numberformat.Format = "dd.MM.yyyy";
                worksheet.Column(10).Style.Numberformat.Format = "0";
                worksheet.Column(14).Style.Numberformat.Format = "0";
                worksheet.Column(27).Style.Numberformat.Format = "0";
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns(minCellSize, maxCellSize);
                package.Save();
            }
        }

    }
}
