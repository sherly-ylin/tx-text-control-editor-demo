namespace editor_demo.Models
{
    public class SNOrder
    {
        public SNOrder()
        {
            OrderID = 0;
            CustomerName = string.Empty;
            BillingAddress = string.Empty;
            DTCreated = DateTime.Now;
            OrderLines = new List<OrderLine>();
        }
        public SNOrder(int orderId, string customerName, string billingAddress, DateTime dtCreated)
        {
            OrderID = orderId;
            CustomerName = customerName;
            BillingAddress = billingAddress;
            DTCreated = dtCreated;
            OrderLines = new List<OrderLine>();
        }
        public int OrderID { get; set; }
        public string CustomerName { get; set; }
        public string BillingAddress { get; set; }
        public DateTime DTCreated { get; set; }
        public List<OrderLine> OrderLines { get; set; } = new List<OrderLine>();
        public decimal TotalSellPrice => OrderLines.Sum(item => item.LineTotal);
        public List<OrderBundle> OrderBundles { get; set; } = new List<OrderBundle>();
        public void GroupOrderLinesByBundle()
        {
            OrderBundles = OrderLines
                .GroupBy(ol => ol.BundleID)
                .Select(g => new OrderBundle
                {
                    BundleID = g.Key,
                    BundleLines = g.ToList(),
                    BundleTotal = g.Sum(ol => ol.LineTotal)
                })
                .ToList();
        }
    }
    public class OrderBundle
    {
        public int BundleID { get; set; }
        public List<OrderLine> BundleLines { get; set; } = new List<OrderLine>();
        public decimal BundleTotal { get; set; }
    }
    public class OrderLine
    {
        public int OrderLineID { get; set; }
        public int BundleID { get; set; }
        public string Model { get; set; }
        public int Quantity { get; set; }
        public decimal SellPrice { get; set; }
        public decimal LineTotal { get; set; }
    }

}
