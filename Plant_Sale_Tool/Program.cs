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
            List<Plant> plants = null;

            if (File.Exists(pathToUse))
            {
                Console.WriteLine(string.Format("Workbook found in {0}", pathToUse));
                using (var package = new ExcelPackage(new FileInfo(pathToUse)))
                {
                    var firstSheet = package.Workbook.Worksheets["Sheet1"];
                    WorkbookHelper workbookHelper = new WorkbookHelper(firstSheet);
                    plants = workbookHelper.ReadWorkBook();
                    foreach (Plant p in plants)
                    {
                        Console.WriteLine(string.Format("{0} = {1} with {2} orders", p.Name, plants.IndexOf(p), p.Orders.Count));
                    }
                }
            } else
            {
                Console.WriteLine(string.Format("Workbook not found in {0}", pathToUse));
                Environment.Exit(1);
            }
    
            bool exit = false;
            do
            {
                Console.WriteLine();
                Console.Write("Enter Plant Name or index that you want more info on:");
                string requestedPlant = Console.ReadLine();
                int inputIndex = 999;
                if (!int.TryParse(requestedPlant, out inputIndex))
                {
                    if(requestedPlant.ToLower() == "q" || requestedPlant.ToLower() == "quit")
                    {
                        break;
                    } else
                    {
                        foreach (Plant p in plants)
                        {
                            if (requestedPlant.ToLower() == p.Name.ToLower())
                            {
                                inputIndex = plants.IndexOf(p);
                                break;
                            }
                            else
                            {
                                inputIndex = 999;
                            }
                        }
                    }
                } 
                
                    
                if (inputIndex == 999)
                {
                    Console.WriteLine("Invalid selection, please make another selection.");
                }
                else
                {
                    Plant selectedPlant = plants[inputIndex];
                    //Console.WriteLine();
                    //Console.WriteLine(string.Format("Plant {0} has {1} orders", selectedPlant.Name, selectedPlant.Orders.Count));
                    //Console.WriteLine(string.Format("Seller\t\t\t\tCustomer\t\t\t\tOrder Amount"));
                    //foreach (Order o in selectedPlant.Orders)
                    //{
                    //    Console.WriteLine(string.Format("{0}\t\t\t\t{1}\t\t\t\t{2}", o.seller, o.customer, o.orderAmount));
                    //}
                    TableHelper tableHelper = new TableHelper(selectedPlant);


                }
                
            } while (!exit);
            


        }
    }
}
