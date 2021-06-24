using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToXslx
{
    class InfoDir
    {
        public static string CsvFolder { get; set; }
        public static string XlsxFolder { get; set; }
        public static  string currentPathFolder = Directory.GetCurrentDirectory();
        //private static string dirName = currentPathFolder + @"\CsvConverter";

        public static void CreateDirectories(string csvFolderName, string xlsxFolderName)
        {
            CsvFolder = currentPathFolder + csvFolderName;
            XlsxFolder = currentPathFolder + xlsxFolderName;
            IsDirectoryExist(currentPathFolder + csvFolderName);
            IsDirectoryExist(currentPathFolder + xlsxFolderName);
        }

        private static void IsDirectoryExist(string directoryName)
        {
            if (!Directory.Exists(directoryName))
            {
                Directory.CreateDirectory(directoryName);
            }
        }

        public static ICollection<string> GetFilesFromDirectory(string dirName) => Directory.GetFiles(dirName, "*.csv").ToList();

    }
}
