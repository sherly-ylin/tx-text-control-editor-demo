namespace editor_demo.Models
{
    public class Order
    {
        public Order()
        {
            CustomerName = string.Empty;
            ShippingAddress = string.Empty;
            OrderDate = DateTime.Now;
            OrderItems = new List<OrderItem>();
        }
        public Order(string customerName, string shippingAddress, DateTime orderDate)
        {
            CustomerName = customerName;
            ShippingAddress = shippingAddress;
            OrderDate = orderDate;
            OrderItems = new List<OrderItem>();
        }
        public string CustomerName { get; set; }
        public string ShippingAddress { get; set; }
        public DateTime OrderDate { get; set; }
        public List<OrderItem> OrderItems { get; set; } = new List<OrderItem>();
        public decimal GrandTotal => OrderItems.Sum(item => item.Total);
    }

    public class OrderItem
    {
        public string ProductName { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Total { get; set; }
    }

}
