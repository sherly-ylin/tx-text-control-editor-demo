using editor_demo.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;
using TXTextControl;
using TXTextControl.DocumentServer;

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

                    // Load the template
                    tx.Load("Documents/template.docx", TXTextControl.StreamType.WordprocessingML, ls);

                    // Convert to format that can be loaded in the editor
                    string documentContent = "";
                    tx.Save(out documentContent, TXTextControl.StringStreamType.HTMLFormat);

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
        // Get available merge fields from the template
        [HttpPost]
        public IActionResult GetMergeFields()
        {
            try
            {
                using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
                {
                    tx.Create();

                    TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings()
                    {
                        ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord
                    };

                    tx.Load("Documents/template.docx", TXTextControl.StreamType.WordprocessingML, ls);

                    // Get all application fields (merge fields)
                    var fields = new List<string>();
                    foreach (TXTextControl.ApplicationField field in tx.ApplicationFields)
                    {
                        fields.Add(field.Name);
                    }

                    return Json(new { success = true, fields = fields });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, error = ex.Message });
            }
        }
        private Order GetSampleData()
        {
            return new Order
            {
                CustomerName = "John Doe",
                ShippingAddress = "123 Main St., Springfield, IL 62701",
                OrderDate = DateTime.Parse("2025-06-13"),

                OrderItems = new List<OrderItem>
                    {
                        new OrderItem
                        {
                            ProductName = "Widget",
                            Quantity = 2,
                            Price = 45.00m,
                            Total = 90.00m
                        },
                        new OrderItem
                        {
                            ProductName = "Gadget",
                            Quantity = 1,
                            Price = 78.45m,
                            Total = 78.45m
                        }
                    }
            };
        }
        static void GenerateOrderDocument(Order order, string templatePath, string outputPath)
        {
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();

                var loadSettings = new LoadSettings
                {
                    ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                    //LoadSubTextParts = true
                };

                tx.Load(templatePath, TXTextControl.StreamType.WordprocessingML, loadSettings);

                using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                {
                    mailMerge.MergeObject(order);
                }

                tx.Save(outputPath, TXTextControl.StreamType.MSWord);
            }
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
