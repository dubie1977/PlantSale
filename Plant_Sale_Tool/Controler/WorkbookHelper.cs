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

        /// <summary>
        /// Reads the workbook
        /// </summary>
        /// <returns></returns>
        public List<Plant> ReadWorkBook()
        {
            List<Plant> Plants = ReadPlants();
            List<Plant> PlantOrders = ReadOrders(Plants);

            return PlantOrders;
        }

        /// <summary>
        /// Reads only the list of plants.
        /// This should be Column A starting with row 2
        /// </summary>
        /// <returns>List of plants</returns>
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

        /// <summary>
        /// Reads the orders for all the plants and adds them to the correct plant.
        /// </summary>
        /// <param name="plants">List of plants</param>
        /// <returns>Fully populated order list</returns>
        private List<Plant> ReadOrders(List<Plant> plants)
        {
            //Read all of the rows in a coloum before moving to the next column
            List<Plant> PlantsWOrders = plants;
            bool stop = false;
            //Loop though all of the orders
            int index = 2;
            do
            {

                if (worksheet.GetValue(2, index) == null)
                {
                    break;
                }
                //Grab seller and customer for all of the orders
                string seller = (string)worksheet.GetValue(1, index);
                string customer = (string)worksheet.GetValue(2, index);

                //Loop though all of the plants in that order and add those orders to the correct plant.
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
            return PlantsWOrders;
        }
    }
}
