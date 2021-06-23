using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace CsvSplitter
{
    public class CsvSplit
    {
        private static int chunkSize;
        public static int ChunkSize
        {
            get
            {
                return chunkSize;
            }
            set
            {
                chunkSize = 240000000;
            }
        }
        public static bool IsLargeFile(string fileName)
        {
            return File.ReadAllLines(fileName).Count() > chunkSize;
        }
        public static void SplitFile(string inputFile, string path)
        {
            const int BUFFER_SIZE = 20 * 1024;
            byte[] buffer = new byte[BUFFER_SIZE];

            using (Stream input = File.OpenRead(inputFile))
            {
                int index = 0;
                while (input.Position < input.Length)
                {
                    using (Stream output = File.Create(path + "\\" + index + ".csv"))
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
