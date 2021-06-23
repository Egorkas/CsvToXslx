﻿using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace SplitterForCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string csvFileName = @"d:\PZ\CsvToXslx\CsvToXslx\bin\Debug\net5.0\1.csv";
            string path = @"d:\PZ\CsvToXslx\CsvToXslx\bin\Debug\net5.0\CsvHelper\";
            SplitFile(csvFileName, 1000000, path);
            Console.WriteLine("Finished");
        }

        public static void SplitFile(string inputFile, int chunkSize, string path)
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
