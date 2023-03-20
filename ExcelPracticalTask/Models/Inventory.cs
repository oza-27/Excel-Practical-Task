namespace ExcelPracticalTask.Models
{
    public class Inventory
    {
        public string ProductCode { get; set; }
        public double EventType { get; set; }
        public double Quantity { get; set; }
        public double Price { get; set; }
        public DateTime Date { get; set; }
        public double TotalPurchaseQuantity { get; set; }
        public double Total_Purchase_Amount { get; set; }
        public double Total_Sale_Quantity { get; set; }
        public double Total_Sale_Amount { get; set; }
        public int Profit { get; set; }
        public int Loss { get; set; }
        public int Opening_Quantity { get; set; }
        public int Closing_Quantity { get; set; }
    }
}
