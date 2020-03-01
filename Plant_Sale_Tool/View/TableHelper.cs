using System;
namespace Plant_Sale_Tool
{
    public class TableHelper
    {
        static int tableWidth = 73;

        public TableHelper() { }

        public static void WriteTable(Plant plant, int tableWidth=73)
        {
            Console.Clear();
            Console.WriteLine(string.Format("Plant {0} has {1} orders", plant.Name, plant.Orders.Count));
            PrintLine();
            PrintRow("Seller", "Customer", "Order Amount");
            PrintLine();
            foreach(Order o in plant.Orders)
            {
                PrintRow(o.seller, o.customer, Convert.ToString(o.orderAmount));
            }
            
            PrintRow("", "", "");
            PrintLine();

        }

        static void PrintLine()
        {
            Console.WriteLine(new string('-', tableWidth));
        }

        static void PrintRow(params string[] columns)
        {
            int width = (tableWidth - columns.Length) / columns.Length;
            string row = "|";

            foreach (string column in columns)
            {
                row += AlignCentre(column, width) + "|";
            }

            Console.WriteLine(row);
        }

        static string AlignCentre(string text, int width)
        {
            text = text.Length > width ? text.Substring(0, width - 3) + "..." : text;

            if (string.IsNullOrEmpty(text))
            {
                return new string(' ', width);
            }
            else
            {
                return text.PadRight(width - (width - text.Length) / 2).PadLeft(width);
            }
        }
    }
}
