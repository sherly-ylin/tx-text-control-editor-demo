namespace editor_demo.Models
{
    public class SNOrder
    {
        public SNOrder()
        {
            CustomerName = string.Empty;
            BillingAddress = string.Empty;
            DTCreated = DateTime.Now;
            OrderLines = new List<OrderLine>();
        }
        public SNOrder(string customerName, string billingAddress, DateTime dtCreated)
        {
            CustomerName = customerName;
            BillingAddress = billingAddress;
            DTCreated = dtCreated;
            OrderLines = new List<OrderLine>();
        }
        public string CustomerName { get; set; }
        public string BillingAddress { get; set; }
        public DateTime DTCreated { get; set; }
        public List<OrderLine> OrderLines { get; set; } = new List<OrderLine>();
        public decimal TotalSellPrice => OrderLines.Sum(item => item.LineTotal);
    }

    public class OrderLine
    {
        public string Model { get; set; }
        public int Quantity { get; set; }
        public decimal SellPrice { get; set; }
        public decimal LineTotal { get; set; }
    }

}
