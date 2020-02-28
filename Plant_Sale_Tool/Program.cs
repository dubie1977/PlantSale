using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Plant_Sale_Tool
{
    class Program
    {
        private const string developemtPath = "../../Plants.xlsx";

        static void Main(string[] args)
        {

            string currentPath = "Plants.xlsx";
            for (int i=0; i < args.Length; i++)
            {                
                Console.WriteLine("arg"+i+" = "+args[i]);
            }
            string fullArgumentsPath = args.Length > 0 ? Path.GetFullPath(args[0]) : "";
            string fullDevelopmentPath = Path.GetFullPath(developemtPath);
            string fullCurrentPath = Path.GetFullPath(currentPath);

            string pathToUse = "";

            if (File.Exists(fullArgumentsPath))
            {
                pathToUse = fullArgumentsPath;
            }
            else if (File.Exists(fullCurrentPath))
            {
                pathToUse = fullCurrentPath;
            } else
            {
                pathToUse = fullDevelopmentPath;
            }

            if (File.Exists(pathToUse))
            {
                Console.WriteLine(string.Format("Workbook found in {0}", pathToUse));
                using (var package = new ExcelPackage(new FileInfo(pathToUse)))
                {
                    var firstSheet = package.Workbook.Worksheets["Sheet1"];
                    WorkbookHelper workbookHelper = new WorkbookHelper(firstSheet);
                    List<Plant> plants = workbookHelper.ReadWorkBook();
                    foreach(Plant p in plants)
                    {
                        Console.WriteLine(string.Format("{0} with {1} orders", p.Name, p.Orders.Count));
                    }
                }
            } else
            {
                Console.WriteLine(string.Format("Workbook not found in {0}", pathToUse));
            }
            
        }
    }
}
