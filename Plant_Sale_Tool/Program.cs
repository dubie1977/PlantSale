using System;
using System.Collections.Generic;
using System.IO;
using DinkToPdf;
using Markdig;
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

            CreateDocumentation();

            
    
            bool exit = false;
            HtmlToPdfDocument pdfDocument = null;
            bool write = false;
            bool writeAll = false;
            do
            {
                Console.WriteLine();
                Console.WriteLine("write or w = write selected plats out to SelectedPlants.pdf");
                Console.WriteLine("all or a = write all plants to AllPlants.pdf");
                Console.WriteLine("quit or q = quit application.");
                Console.Write("Enter Plant Name or index that you want more info on:");
                string userInput = Console.ReadLine();
                int inputIndex = 999;
                if (!int.TryParse(userInput, out inputIndex))
                {
                    if (userInput.ToLower() == "q" || userInput.ToLower() == "quit")
                    {
                        break;
                    } else if ((userInput.ToLower() == "w" || userInput.ToLower() == "write") && pdfDocument != null)
                    {
                        write = true;
                    }
                    else if (userInput.ToLower() == "a" || userInput.ToLower() == "all")
                    {
                        writeAll = true;
                    }
                    else
                    {
                        foreach (Plant p in plants)
                        {
                            if (userInput.ToLower() == p.Name.ToLower())
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


                if (write)
                {
                    File.WriteAllBytes("SelectedPlants.pdf", PdfCreatorController.ConvertDocument(pdfDocument));
                    pdfDocument = null;
                    write = false;
                }
                else if (writeAll)
                {
                    PdfCreatorController pdfCreator = new PdfCreatorController();
                    foreach (Plant p in plants)
                    {
                        string html = HTMLGenerator.GetHTMLString(p);
                        pdfDocument = pdfCreator.AddPageToPDF(html, p.Name, pdfDocument);
                    }

                    File.WriteAllBytes("AllPlants.pdf", PdfCreatorController.ConvertDocument(pdfDocument));
                    pdfDocument = null;
                    writeAll = false;
                }
                else if (inputIndex == 999)
                {
                    Console.WriteLine("Invalid selection, please make another selection.");
                }
                else
                {
                    Plant selectedPlant = plants[inputIndex];
                    TableHelper.WriteTable(selectedPlant);
                    string html = HTMLGenerator.GetHTMLString(selectedPlant);
                    Console.WriteLine();
                    PdfCreatorController pdfCreator = new PdfCreatorController();
                    pdfDocument = pdfCreator.AddPageToPDF(html, selectedPlant.Name, pdfDocument);


                }


                
            } while (!exit);
            


        }

        public static byte[] Combine(byte[] first, byte[] second)
        {
            if(first == null && second != null)
            {
                return second;
            } else if(first != null & second == null)
            {
                return first;
            } else if(first == null && second == null)
            {
                return null;
            }

            byte[] bytes = new byte[first.Length + second.Length];
            Buffer.BlockCopy(first, 0, bytes, 0, first.Length);
            Buffer.BlockCopy(second, 0, bytes, first.Length, second.Length);
            return bytes;
        }

        private static void CreateDocumentation()
        {
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            var html = Markdown.ToHtml("This is a text with some *emphasis*  ", pipeline);

            PdfCreatorController pdfCreator = new PdfCreatorController();
            HtmlToPdfDocument pdfDocument = pdfCreator.AddPageToPDF(html, "Plante Sale Guide");
            File.WriteAllBytes("Plant Sale Guide.pdf", PdfCreatorController.ConvertDocument(pdfDocument));
        }
    }
}
