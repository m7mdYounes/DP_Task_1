using CsvHelper;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using System.Diagnostics;
using Task_1.Models;
using CsvHelper.Configuration;
using System.Globalization;
namespace Task_1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        public IActionResult index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var receipts = new List<Receipts>();
                try
                {
                    using (var reader = new StreamReader(file.OpenReadStream()))
                    using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        HasHeaderRecord = true
                    }))
                    {
                        csv.Context.RegisterClassMap<ReceiptsMap>();
                        csv.Read();
                        csv.ReadHeader();

                        receipts = csv.GetRecords<Receipts>().ToList();
                    }

                    TempData["Receipts"] = Newtonsoft.Json.JsonConvert.SerializeObject(receipts);

                    return View("GridView", receipts);
                }
                catch (Exception ex)
                {
                
                    ModelState.AddModelError(string.Empty, "An error occurred while processing the file: " + ex.Message);
                    return View("Index");
                }
            }

            ModelState.AddModelError(string.Empty, "Please upload a valid CSV file.");
            return RedirectToAction("Index");

        }
    
        
        public IActionResult DownloadXlsx()
        {
            var receipts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Receipts>>(TempData["Receipts"] as string);
            if (receipts == null)
            {
                return RedirectToAction("Index");
            }

            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Receipts");

                worksheet.Cells[1, 1].Value = "Business Unit";
                worksheet.Cells[1, 2].Value = "Receipt Method Id";
                worksheet.Cells[1, 3].Value = "Remittance Bank";
                worksheet.Cells[1, 4].Value = "Remittance Bank Account";
                worksheet.Cells[1, 5].Value = "Receipt Number";
                worksheet.Cells[1, 6].Value = "Receipt Amount";
                worksheet.Cells[1, 7].Value = "Receipt Date";
                worksheet.Cells[1, 8].Value = "Accounting Date";
                worksheet.Cells[1, 9].Value = "Conversion Date";
                worksheet.Cells[1, 10].Value = "Currency";
                worksheet.Cells[1, 11].Value = "Conversion Rate Type";
                worksheet.Cells[1, 12].Value = "Conversion Rate";
                worksheet.Cells[1, 13].Value = "Customer Name";
                worksheet.Cells[1, 14].Value = "Customer Account Number";
                worksheet.Cells[1, 15].Value = "Customer Site Number";
                worksheet.Cells[1, 16].Value = "Invoice Number Reference";
                worksheet.Cells[1, 17].Value = "Invoice Amount";
                worksheet.Cells[1, 18].Value = "Comments";


                for (int i = 0; i < receipts.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = receipts[i].businessUnit;
                    worksheet.Cells[i + 2, 2].Value = receipts[i].receiptMethodId;
                    worksheet.Cells[i + 2, 3].Value = receipts[i].remittanceBank;
                    worksheet.Cells[i + 2, 4].Value = receipts[i].remittanceBankAccount;
                    worksheet.Cells[i + 2, 5].Value = receipts[i].receiptNumber;
                    worksheet.Cells[i + 2, 6].Value = receipts[i].receiptAmount;
                    worksheet.Cells[i + 2, 7].Value = receipts[i].receiptDate;
                    worksheet.Cells[i + 2, 8].Value = receipts[i].accountingDate;
                    worksheet.Cells[i + 2, 9].Value = receipts[i].conversionDate;
                    worksheet.Cells[i + 2, 10].Value = receipts[i].currency;
                    worksheet.Cells[i + 2, 11].Value = receipts[i].conversionRateType;
                    worksheet.Cells[i + 2, 12].Value = receipts[i].conversionRate;
                    worksheet.Cells[i + 2, 13].Value = receipts[i].customerName;
                    worksheet.Cells[i + 2, 14].Value = receipts[i].customerAccountNumber;
                    worksheet.Cells[i + 2, 15].Value = receipts[i].customerSiteNumber;
                    worksheet.Cells[i + 2, 16].Value = receipts[i].invoiceNumberReference;
                    worksheet.Cells[i + 2, 17].Value = receipts[i].invoiceAmount;
                    worksheet.Cells[i + 2, 18].Value = receipts[i].comments;

                }
                //worksheet.Cells.LoadFromCollection(receipts, true);
                package.Save();

            }
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Receipts.xlsx");
        }







        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
