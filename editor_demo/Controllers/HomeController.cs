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
                using (ServerTextControl tx = new ServerTextControl())
                {
                    tx.Create();

                    LoadSettings ls = new LoadSettings()
                    {
                        ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                        LoadSubTextParts = true
                    };

                    tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, ls);

                    // string mergedDoc = "";
                    // mergedDoc = MergeData(tx);
                    // // Pass the document content to the view
                    // ViewBag.TemplateDocument = mergedDoc;
                    // ViewBag.TemplateLoaded = true;

                    // Convert to format that can be loaded in the editor
                    string documentContent = "";
                    tx.Save(out documentContent, StringStreamType.HTMLFormat);
                    // Pass the document content to the view
                    ViewBag.TemplateDocument = documentContent;
                    ViewBag.TemplateLoaded = true;

                    // // save document as PDF
                    // byte[] document;
                    // tx.Save(out document, TXTextControl.BinaryStreamType.AdobePDF);
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
                using (ServerTextControl tx = new ServerTextControl())
                {
                    tx.Create();

                    // Get the HTML content from the editor
                    string htmlContent = Request.Form["textcontrol"];

                    if (string.IsNullOrEmpty(htmlContent))
                    {
                        return Json(new { success = false, error = "No content to save" });
                    }

                    // Load the HTML content into ServerTextControl
                    tx.Load(htmlContent, StringStreamType.HTMLFormat);

                    // Save back to the original template file
                    tx.Save("Documents/output.docx", StreamType.WordprocessingML);

                    return Json(new { success = true, message = "Document saved successfully!" });
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
                                WHERE o.OrderId = @OrderId
                                ORDER BY ol.BundleID, Model";
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(query, conn);

                    cmd.Parameters.AddWithValue("@OrderId", orderId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        resultTable.Load(reader);
                        Console.WriteLine("Total results rows: " + resultTable.Rows.Count);
                    }
                    if (resultTable.Rows.Count > 0)
                    {
                        var row = resultTable.Rows[0];
                        order.OrderID = Convert.ToInt32(row["OrderID"]);
                        order.CustomerName = row["CustomerName"].ToString();
                        order.BillingAddress = row["BillingAddress1"].ToString() + ", " +
                            (string.IsNullOrEmpty(row["BillingAddress2"].ToString()) ? "" : row["BillingAddress2"].ToString() + ", ") +
                            row["BillingCity"].ToString() + ", " +
                            row["BillingState"].ToString() + " " +
                            row["BillingPostalCode"].ToString();
                        order.DTCreated = Convert.ToDateTime(row["DTCreated"]);

                        // Loop through all rows to get order items
                        foreach (DataRow itemRow in resultTable.Rows)
                        {
                            order.OrderLines.Add(new OrderLine
                            {
                                OrderLineID = Convert.ToInt32(itemRow["OrderLineID"]),
                                BundleID = Convert.ToInt32(itemRow["BundleID"]),
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

                finally
                {
                    conn.Close();
                }
            }

            return order;
        }


        // private Order GetSampleData()
        // {
        //     return new Order
        //     {
        //         CustomerName = "John Doe",
        //         ShippingAddress = "123 Main St., Springfield, IL 62701",
        //         OrderDate = DateTime.Parse("2025-06-13"),

        //         OrderItems = new List<OrderItem>
        //             {
        //                 new OrderItem
        //                 {
        //                     ProductName = "Widget",
        //                     Quantity = 2,
        //                     Price = 45.00m,
        //                     Total = 90.00m
        //                 },
        //                 new OrderItem
        //                 {
        //                     ProductName = "Gadget",
        //                     Quantity = 1,
        //                     Price = 78.45m,
        //                     Total = 78.45m
        //                 }
        //             }
        //     };
        // }
        public string MergeData(ServerTextControl tx)
        {
            using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
            {
                SNOrder dbOrder = GetOrderFromDb(7262);
                var dbOrders = new List<SNOrder> { dbOrder };
                Console.WriteLine(JsonConvert.SerializeObject(dbOrder));

                mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                mailMerge.MergeObject(dbOrder);
            }
            // Convert to format that can be loaded in the editor
            string documentContent = "";
            tx.Save(out documentContent, StringStreamType.HTMLFormat);


            return documentContent;
        }

        [HttpPost]
        public IActionResult MergeOrderData(int orderId)
        {
            try
            {
                using (ServerTextControl tx = new ServerTextControl())
                {
                    tx.Create();

                    // Load the template
                    var loadSettings = new LoadSettings
                    {
                        ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                        LoadSubTextParts = true
                    };
                    tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, loadSettings);

                    // Merge the data
                    SNOrder dbOrder = GetOrderFromDb(orderId);
                    using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                    {
                        mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                        mailMerge.MergeObject(dbOrder);
                    }

                    // Return merged HTML to update editor
                    string mergedHtml = "";
                    tx.Save(out mergedHtml, StringStreamType.HTMLFormat);
                    return Json(new { success = true, mergedHtml });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, error = ex.Message });
            }
        }


        [HttpPost]
        public IActionResult GenerateDBOrderDocument(int orderId, string templatePath, string outputPath)
        {
            try
            {
                // Get the order from the database
                SNOrder dbOrder = GetOrderFromDb(orderId);
                Console.WriteLine(JsonConvert.SerializeObject(dbOrder));

                // Define paths for template and output document
                templatePath = "Documents/template_order.docx";
                outputPath = $"Documents/SNOrder_{orderId}.docx";

                using (ServerTextControl tx = new ServerTextControl())
                {
                    tx.Create();

                    // Load the template document
                    var loadSettings = new LoadSettings
                    {
                        ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                    };
                    tx.Load(templatePath, StreamType.WordprocessingML, loadSettings);

                    // Perform mail merge with the SNOrder object
                    using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                    {
                        mailMerge.MergeObject(dbOrder);
                    }
                    // Return merged document
                    string mergedDocument = "";
                    tx.Save(out mergedDocument, StringStreamType.HTMLFormat);
                    ViewBag.TemplateDocument = mergedDocument;
                    // Save the generated document
                    tx.Save(outputPath, StreamType.WordprocessingML);
                    return Json(new { success = true, message = "Document generated successfully!", filePath = outputPath });

                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, error = ex.Message });
            }
        }
        static void DownloadDocument(ServerTextControl tx, string outputPath)
        {
            tx.Save(outputPath, StreamType.MSWord);
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
