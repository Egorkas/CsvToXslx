using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using CsvSplitter;


namespace CsvToXslx
{
    class Program
    {
        public static void Main()
        {
            //for support rus ANSI in .NET Core
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string csvFolderName = @"\CsvConverter";
            string xlsxFolderName = @"\XlsxResult";
            InfoDir.CreateDirectories(csvFolderName, xlsxFolderName);

            LargeFileMonitor(InfoDir.CsvFolder);

            FilesConverter(InfoDir.CsvFolder, InfoDir.XlsxFolder);

            Console.WriteLine("Finished!");
        }

        public static void FilesConverter(string inputFolder, string outputFolder)
        {
            var files = InfoDir.GetFilesFromDirectory(inputFolder);
            foreach (var item in files)
            {
                CsvConverterToXlsx.CsvConvert(item, outputFolder);
            }
        }
        public static void LargeFileMonitor(string folder)
        {
            var files = InfoDir.GetFilesFromDirectory(InfoDir.CsvFolder);
            foreach (var item in files)
            {
                if (CsvSplit.IsManyLinesFile(item))
                {
                    CsvSplit.SplitFile(item, InfoDir.CsvFolder);
                    File.Delete(item);
                }
            }
        }
    }
}


