using System;
using System.IO;
using OfficeOpenXml;

namespace Plant_Sale_Tool
{
    class Program
    {
        static void Main(string[] args)
        {

            DirectoryInfo info = new DirectoryInfo("../../../../");
            using (var package = new ExcelPackage(new FileInfo("../../../../Book.xlsx")))
            {
                var firstSheet = package.Workbook.Worksheets["Sheet1"];
                Console.WriteLine("Sheet 1 Data");
                Console.WriteLine($"Cell A1 Value   : {firstSheet.Cells["A1"].Text}");
                bool stop = false;
                int index = 3;
                do
                {
                    string value = firstSheet.Cells["A" + index.ToString()].Text;
                    Console.WriteLine( string.Format("Cell A{0} = {1}", index, value) );
                    stop = string.IsNullOrWhiteSpace(value);
                    index++;
                } while (!stop);
                //Console.WriteLine($"Cell A2 Color   : {firstSheet.Cells["A2"].Style.Font.Color.LookupColor()}");
                //Console.WriteLine($"Cell B2 Formula : {firstSheet.Cells["B2"].Formula}");
                //Console.WriteLine($"Cell B2 Value   : {firstSheet.Cells["B2"].Text}");
                //Console.WriteLine($"Cell B2 Border  : {firstSheet.Cells["B2"].Style.Border.Top.Style}");
                //Console.WriteLine("");

                //var secondSheet = package.Workbook.Worksheets["Sheet2"];
                //Console.WriteLine($"Sheet 2 Data");
                //Console.WriteLine($"Cell A2 Formula : {secondSheet.Cells["A2"].Formula}");
                //Console.WriteLine($"Cell A2 Value   : {secondSheet.Cells["A2"].Text}");
            }
        }
    }
}
