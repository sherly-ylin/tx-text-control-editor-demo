using System.Diagnostics;
using editor_demo.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json; 

namespace editor_demo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {

            //using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
            //{
            //    tx.Create();
            //    // Load your RTF template
            //    TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings()
            //    {
            //        ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord
            //    };

            //    // Load the RTF template file
            //    tx.Load("Documents/order.rtf", TXTextControl.StreamType.RichTextFormat, ls);

            //    // Prepare sample data for merge
            //    var orderData = new
            //    {
            //        CustomerName = "SalesChain LLC",
            //        OrderDate = DateTime.Now.ToString("yyyy-MM-dd"),
            //        ShippingAddress = "123 Business Street, Suite 100, City, State 12345",
            //        OrderItems = new[]
            //        {
            //            new { ProductName = "Laptop Computer", Quantity = 2, Price = 999.99, Total = 1999.98 },
            //            new { ProductName = "Wireless Mouse", Quantity = 2, Price = 29.99, Total = 59.98 },
            //            new { ProductName = "USB Cable", Quantity = 3, Price = 15.99, Total = 47.97 }
            //        },
            //        GrandTotal = 2107.93
            //    };

            //    // Convert to JSON for MailMerge
            //    string jsonData = JsonConvert.SerializeObject(new[] { orderData });

            //    // Perform mail merge
            //    using (TXTextControl.DocumentServer.MailMerge mailMerge = new TXTextControl.DocumentServer.MailMerge())
            //    {
            //        mailMerge.TextComponent = tx;

            //        // Merge the JSON data
            //        mailMerge.MergeJsonData(jsonData);
            //    }

            //    // Get the merged result as HTML
            //    string result = "";
            //    tx.Save(out result, TXTextControl.StringStreamType.HTMLFormat);

            //    ViewBag.Document = result;
            //}

            try
            {
                using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
                {
                    tx.Create();

                    // Load settings for merge fields
                    TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings()
                    {
                        ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord
                    };

                    // Load the RTF template
                    tx.Load("Documents/order.rtf", TXTextControl.StreamType.RichTextFormat, ls);

                    // Convert to format that can be loaded in the editor
                    string documentContent = "";
                    tx.Save(out documentContent, TXTextControl.StringStreamType.RichTextFormat);

                    // Pass the document content to the view
                    ViewBag.TemplateDocument = documentContent;
                    ViewBag.TemplateLoaded = true;
                }
            }
            catch (Exception ex)
            {
                ViewBag.TemplateLoaded = false;
                ViewBag.ErrorMessage = ex.Message;
            }
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
