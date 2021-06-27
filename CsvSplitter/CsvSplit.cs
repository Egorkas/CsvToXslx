using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace CsvSplitter
{
    public class CsvSplit
    {
        private static int chunkSize = 240000000;//Mb
        private static int linesCount = 1000000;//Count of lines, because for  EPPluS MaxRows = 1048576;
        public static int FileLinesCount { get; set; } 
        
        public static int ChunkSize
        {
            get
            {
                return chunkSize;
            }
            set
            {
                chunkSize = value;
            }
        }
        public static bool IsManyLinesFile(string fileName)
        {
            FileLinesCount = File.ReadAllLines(fileName).Count();
            return FileLinesCount > linesCount;
        }
        public static bool IsLargeFile(string fileName)
        {
            FileInfo memorySize = new FileInfo(fileName);
            return memorySize.Length > chunkSize;
        }

        public static void SplitFile(string inputPath, string outputPath)
        {
            int countOutputFiles = FileLinesCount % linesCount == 0 ? FileLinesCount / linesCount : FileLinesCount / linesCount + 1;
            string fileName = Path.GetFileNameWithoutExtension(inputPath);
            using(StreamReader sr = new StreamReader(inputPath, Encoding.GetEncoding(1251)))
            {
                int fileNumber = 1;
                while (!sr.EndOfStream && fileNumber <= countOutputFiles)
                {
                    using(StreamWriter sw = new StreamWriter(outputPath + "\\" + fileName + "_" + fileNumber + ".csv", false, Encoding.GetEncoding(1251)))
                    {
                        if (fileNumber == countOutputFiles)
                        {
                            sw.Write(sr.ReadToEnd());
                            break;
                        }
                        for (int i = 0; i < linesCount; i++)
                        {
                            sw.WriteLine(sr.ReadLine());
                        }
                        fileNumber++;
                    }
                }
            }                      
        }

        public static void SplitFileForMemorySize(string inputPath, string outputPath)
        {
            const int BUFFER_SIZE = 20 * 1024;
            byte[] buffer = new byte[BUFFER_SIZE];
            string fileName = Path.GetFileNameWithoutExtension(inputPath);
            using (Stream input = File.OpenRead(inputPath))
            {
                int index = 0;
                while (input.Position < input.Length)
                {
                    using (Stream output = File.Create(outputPath + "\\" + fileName + "_" + index + ".csv"))
                    {
                        int remaining = chunkSize, bytesRead;
                        while (remaining > 0 && (bytesRead = input.Read(buffer, 0,
                                Math.Min(remaining, BUFFER_SIZE))) > 0)
                        {
                            output.Write(buffer, 0, bytesRead);
                            remaining -= bytesRead;
                        }
                    }
                    index++;
                    Thread.Sleep(500); // experimental; perhaps try it
                }
            }
        }

    }
}
