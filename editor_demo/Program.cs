using editor_demo.Models;
using TXTextControl;
using TXTextControl.DocumentServer;
using TXTextControl.Web;

// Create order object
Order order = CreateOrder();

// Process document generation
GenerateOrderDocument(order, "Documents/template.docx", "Documents/output.pdf");

static Order CreateOrder()
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

        tx.Save(outputPath, TXTextControl.StreamType.AdobePDF);
    }
}

//var builder = WebApplication.CreateBuilder(args);

//// Add services to the container.
//builder.Services.AddControllersWithViews();

//var app = builder.Build();

//// Configure the HTTP request pipeline.
//if (!app.Environment.IsDevelopment())
//{
//    app.UseExceptionHandler("/Home/Error");
//    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
//    app.UseHsts();
//}

//app.UseHttpsRedirection();
//app.UseStaticFiles();

//// enable Web Sockets
//app.UseWebSockets();

//// attach the Text Control WebSocketHandler middleware
//app.UseTXWebSocketMiddleware();

//app.UseRouting();

//app.UseAuthorization();

//app.MapControllerRoute(
//    name: "default",
//    pattern: "{controller=Home}/{action=Index}/{id?}"); //Controller name and method name(also same as view file name)

//app.Run();
