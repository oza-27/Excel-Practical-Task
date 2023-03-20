using ExcelDataReader;
using ExcelPracticalTask.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.POIFS.Crypt.Dsig;

namespace ExcelPracticalTask.Controllers
{
    public class FileController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        // Reading an Excel File
        public IActionResult ReadExcel()
        {
            var prodCod = "";
            var totalPurQnt = "";
            double totalPurAmt = 0;
            var totalSaleQnt = "";
            double totalSaleAmt = 0;
            double profitLoss = 0;

            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "inevtory task.xlsx");

            // creating a list to store ExcelData
            var excelData = new List<Inventory>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            // opening the excel file using package ExcelDataReader
            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        var rowData = new Inventory();

                        // Get the cell values for each columns

                        rowData.ProductCode = reader.GetString(0);
                        rowData.EventType = reader.GetDouble(1);
                        rowData.Quantity = reader.GetDouble(2);
                        rowData.Price = reader.GetDouble(3);
                        rowData.Date = reader.GetDateTime(4);

                        // Adding a row to excel data
                        excelData.Add(rowData);
                    }
                }
            }

            // groupby month-wise
            var monthWise = excelData.GroupBy(x => x.Date.Month).ToList();

            foreach (var month in monthWise)
            {
                var product = month.GroupBy(x => x.ProductCode).ToList();

                foreach (var type in product)
                {
                    foreach (var item in type)
                    {
                        if (item.EventType == 1)
                        {
                            totalPurQnt = item.Quantity.ToString();
                            totalPurAmt = item.Price;
                            var total = Convert.ToDouble(totalPurQnt) * totalPurAmt;
                        }
                        else
                        {
                            totalSaleQnt = item.Quantity.ToString();
                            totalSaleAmt = item.Price;
                            var totalSale = Convert.ToDouble(totalSaleQnt) * totalSaleAmt;
                        }
                    }
                }
            }
            return View(excelData);
        }
    }
}

