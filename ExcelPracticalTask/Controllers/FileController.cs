﻿using ExcelDataReader;
using ExcelPracticalTask.Models;
using MathNet.Numerics;
using Microsoft.AspNetCore.Mvc;
using NPOI.POIFS.Crypt.Dsig;
using System;

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
            double totalPurQnt = 0;
            double totalPurAmt = 0;
            double totalSaleQnt = 0;
            double totalSaleAmt = 0;
            double profitLoss = 0;
            double totalPurcharse = 0;
            double totalSale = 0;
            double profitloss = 0;


            List<InventoryVM> inventoryVM = new List<InventoryVM>();

            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Book1.xlsx");

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
                    var date = DateTime.Now;
                    var productCode = "";
                    totalSaleQnt = 0;
                    prodCod = "";
                    totalPurQnt = 0;
                    totalPurAmt = 0;
                    totalSaleQnt = 0;
                    totalSaleAmt = 0;
                    profitLoss = 0;
                    totalPurcharse = 0;
                    totalSale = 0;
                    double closeQty = 0;
                    double openQty = 0;

                    foreach (var item in type)
                    {
                        date = item.Date;
                        productCode = item.ProductCode;
                        if (item.EventType == 1)
                        {
                            totalPurQnt += item.Quantity;
                            totalPurAmt += item.Price;
                            totalPurcharse = totalPurcharse + (Convert.ToDouble(item.Price) * item.Quantity);
                        }
                        else if (item.EventType == 2)
                        {
                            totalSaleQnt = item.Quantity;
                            totalSaleAmt = item.Price;
                            totalSale = Convert.ToDouble(totalSaleQnt) * totalSaleAmt;
                          
                        }
                        profitloss = (totalSaleAmt * totalSaleQnt) - (totalSaleQnt * totalPurAmt);

                    }
                    var resultOpen = inventoryVM.FirstOrDefault(x => x.ProductCode == productCode.ToString());

                    if (resultOpen != null)
                    {
                        inventoryVM.Add(new InventoryVM
                        {
                            Date = date,
                            ProductCode = productCode,
                            TotalPurchaseQuantity = Convert.ToDouble(totalPurQnt),
                            Total_Purchase_Amount = totalPurcharse.ToString(),
                            Total_Sale_Quantity = totalSaleQnt.ToString(),
                            Total_Sale_Amount = Convert.ToString(totalSale),
                            Profit_Loss = profitloss.ToString(),
                            Closing_Quantity = (resultOpen.Closing_Quantity) + totalPurQnt - totalSaleQnt,
                            Opening_Quantity = resultOpen.Closing_Quantity
                        });
                    }
                    else
                    {

                        inventoryVM.Add(new InventoryVM
                        {
                            Date = date,
                            ProductCode = productCode,
                            TotalPurchaseQuantity = Convert.ToDouble(totalPurQnt),
                            Total_Purchase_Amount = totalPurcharse.ToString(),
                            Total_Sale_Quantity = totalSaleQnt.ToString(),
                            Total_Sale_Amount = Convert.ToString(totalSale),
                            Profit_Loss = profitloss.ToString(),
                            Closing_Quantity = (totalPurQnt - totalSaleQnt),
                            Opening_Quantity = 0,
                        });
                    }
                }
            }
            return View(inventoryVM);
        }
    }
}

