using System;
using System.Collections.Generic;

namespace Plant_Sale_Tool
{
    public class Plant
    {
        public string Name { get; set; }
        public List<Order> Orders = new List<Order>();

        public Plant()
        {
        }

        public Plant(string name)
        { 
            Name = name;
        }
    }
}
