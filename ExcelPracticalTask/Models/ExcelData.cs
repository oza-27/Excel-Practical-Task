namespace ExcelPracticalTask.Models
{
    public class ExcelData
    {
        public DateOnly DataYear { get; set; }
        public string ProductCode { get; set; }
        public int EventType { get; set; }
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
