namespace ExcelPracticalTask.Models
{
    public class InventoryVM
    {
        public DateTime Date { get; set; }
        public string ProductCode { get; set; }
        public double EventType { get; set; }
        public double TotalPurchaseQuantity { get; set; }
        public string Total_Purchase_Amount { get; set; }
        public string Total_Sale_Quantity { get; set; }
        public string Total_Sale_Amount { get; set; }
        public string Profit_Loss { get; set; }
        public double Opening_Quantity { get; set; }
        public double Closing_Quantity { get; set; }
    }
}
