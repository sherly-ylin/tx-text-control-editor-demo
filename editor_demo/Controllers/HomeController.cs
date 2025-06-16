using editor_demo.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Data;
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

                    TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings()
                    {
                        ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord
                    };

                    tx.Load("Documents/template.docx", TXTextControl.StreamType.WordprocessingML, ls);

                    SNOrder dbOrder = GetOrderFromDb(7262);
                    Console.WriteLine(JsonConvert.SerializeObject(dbOrder));

                    using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                    {
                        Order order = GetSampleData();
                        var orders = new List<Order> { order };
                        mailMerge.FormFieldMergeType = FormFieldMergeType.Replace;
                        mailMerge.MergeObjects(orders);
                    }

                    // Convert to format that can be loaded in the editor
                    string documentContent = "";
                    tx.Save(out documentContent, TXTextControl.StringStreamType.HTMLFormat);

                    // Pass the document content to the view
                    ViewBag.TemplateDocument = documentContent;
                    ViewBag.TemplateLoaded = true;

                    // save document as PDF
                    byte[] document;
                    tx.Save(out document, TXTextControl.BinaryStreamType.AdobePDF);
                }
            }
            catch (Exception ex)
            {
                ViewBag.TemplateLoaded = false;
                ViewBag.ErrorMessage = ex.Message;
            }
            return View();
        }

        // Save the document back to the original file
        [HttpPost]
        public IActionResult SaveDocument()
        {
            try
            {
                using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
                {
                    tx.Create();

                    // Get the HTML content from the editor
                    string htmlContent = Request.Form["textcontrol"];

                    if (string.IsNullOrEmpty(htmlContent))
                    {
                        return Json(new { success = false, error = "No content to save" });
                    }

                    // Load the HTML content into ServerTextControl
                    tx.Load(htmlContent, TXTextControl.StringStreamType.HTMLFormat);

                    // Save back to the original template file
                    tx.Save("Documents/output.docx", TXTextControl.StreamType.WordprocessingML);

                    return Json(new { success = true, message = "Document saved successfully!" });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, error = ex.Message });
            }
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

                    tx.Load("Documents/template2.docx", TXTextControl.StreamType.WordprocessingML, ls);

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
        public SNOrder GetOrderFromDb(int orderId)
        {
            var order = new SNOrder();

            string connectionString = "Server=192.168.20.97;Database=SalesChain0602_MS_MN;User Id=ylin;Password=9244@Wahg;TrustServerCertificate=True;";
            DataTable resultTable = new DataTable();
            
            using (var conn = new Microsoft.Data.SqlClient.SqlConnection(connectionString))
            {
                conn.Open();
                try
                {
                    var query = @"SELECT * FROM SNOrder o 
                                JOIN SNOrderLine ol on o.OrderId = ol.OrderId
                                WHERE o.OrderId = @OrderId";
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(query, conn);

                    cmd.Parameters.AddWithValue("@OrderId", orderId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        resultTable.Load(reader);
                        Console.WriteLine(resultTable.Rows.Count);
                    }
                    if (resultTable.Rows.Count > 0)
                    {
                        // Assuming the first row contains the order info
                        var row = resultTable.Rows[0];
                        order.CustomerName = row["CustomerName"].ToString();
                        order.BillingAddress = row["BillingAddress1"].ToString() + ", " +
                                              row["BillingAddress2"].ToString() + ", " +
                                              row["BillingCity"].ToString() + ", " +
                                              row["BillingState"].ToString() + " " +
                                              row["BillingPostalCode"].ToString();
                        order.DTCreated = Convert.ToDateTime(row["DTCreated"]);

                        // Loop through all rows to get order items
                        foreach (DataRow itemRow in resultTable.Rows)
                        {
                            order.OrderLines.Add(new OrderLine
                            {
                                Model = itemRow["Model"].ToString(),
                                Quantity = Convert.ToInt32(itemRow["Quantity"]),
                                SellPrice = Convert.ToDecimal(itemRow["SellPrice"]),
                                LineTotal = Convert.ToDecimal(itemRow["LineTotal"])
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle exception (log or rethrow as needed)
                    throw new Exception("Error retrieving order info: " + ex.Message, ex);
                }

                // // Get order items
                // using (var cmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT ProductName, Quantity, Price, Total FROM OrderItems WHERE OrderId = @OrderId", conn))
                // {
                //     cmd.Parameters.AddWithValue("@OrderId", orderId);
                //     using (var reader = cmd.ExecuteReader())
                //     {
                //         while (reader.Read())
                //         {
                //             order.OrderItems.Add(new OrderItem
                //             {
                //                 ProductName = reader.GetString(0),
                //                 Quantity = reader.GetInt32(1),
                //                 Price = reader.GetDecimal(2),
                //                 Total = reader.GetDecimal(3)
                //             });
                //         }
                //     }
                // }
            }

            return order;
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
