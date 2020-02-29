using System;
namespace Plant_Sale_Tool
{
    public class Order
    {
        public string seller { get; set; }
        public string customer { get; set; }
        public double orderAmount { get; set; }

        public Order()
        {

        }

        public Order(string seller, string customer, double orderAmount)
        {
            this.seller = seller != null ? this.seller = seller : this.seller = "Unknown";
            this.customer = customer;
            this.orderAmount = orderAmount;
        }
    }
}
