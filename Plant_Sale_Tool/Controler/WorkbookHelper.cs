using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace Plant_Sale_Tool
{
    public class WorkbookHelper
    {

        private ExcelWorksheet worksheet;
        public WorkbookHelper(ExcelWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public List<Plant> ReadWorkBook()
        {
            List<Plant> Plants = ReadPlants();
            List<Plant> PlantOrders = ReadOrders(Plants);

            return PlantOrders;
        }

        private List<Plant> ReadPlants()
        {
            List<Plant> plants = new List<Plant>();
            Console.WriteLine("Sheet 1 Data");
            Console.WriteLine($"Cell A1 Value   : {worksheet.Cells["A1"].Text}");
            bool stop = false;
            int index = 3;
            do
            {
                string value = worksheet.Cells["A" + index.ToString()].Text;
                if (string.IsNullOrWhiteSpace(value))
                {
                    break;
                }
                Plant plant = new Plant(value);
                //Console.WriteLine(string.Format("Cell A{0} = {1}", index, value));

                plants.Add(plant);
                index++;
            } while (!stop);


            return plants;
        }

        private List<Plant> ReadOrders(List<Plant> plants)
        {
            List<Plant> PlantsWOrders = plants;
            bool stop = false;
            int index = 2;
            do
            {
                //string value = worksheet.Cells["A" + index.ToString()].Text;
                

                if (worksheet.GetValue(2, index) == null)
                {
                    break;
                }
                string seller = (string)worksheet.GetValue(1, index);
                string customer = (string)worksheet.GetValue(2, index);
                //Console.WriteLine(string.Format("Cell A{0} = {1}", index, value));

                int pIndex = 3;
                do
                {
                    object amount = worksheet.GetValue(pIndex, index);
                    if (amount != null)
                    {
                        Order order = new Order(seller, customer, (double)amount);
                        PlantsWOrders[pIndex - 3].Orders.Add(order);
                    }
                    pIndex++;
                } while (pIndex - 3 <= plants.Count);
                index++;
            } while (!stop);
            Order newOrder = new Order
                (
                    (string)worksheet.GetValue(1, 7),
                    (string)worksheet.GetValue(2, 7),
                    (double)worksheet.GetValue(3, 7)
                );
            PlantsWOrders[0].Orders.Add(newOrder);

            return PlantsWOrders;
        }
    }
}
