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
            
            var files = InfoDir.GetFilesFromDirectory(InfoDir.CsvFolder);

            foreach (var item in files)
            {
                if(CsvSplit.IsLargeFile(item))
                {
                    CsvSplit.SplitFile(item, InfoDir.CsvFolder);

                }
                    

            }
            

            Console.WriteLine("Finished!");
        }
    }
}


