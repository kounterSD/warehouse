using OfficeOpenXml;


namespace Warehouse;

public class Excel
{
    //list gets loaded from Excel file -->in start of the Main method.
    List<Product> products = new List<Product>();
    string path = "./../../../warehouse.xlsx";
    Hash hash = new();
    
    //Returns Catalog in a list Format.
    public List<Product> ReadCatalog(string filename)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filename)))
        {
            var worksheet = package.Workbook.Worksheets["warehouse"];
            worksheet.Cells[1,4].Value = "Status";
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                products.Clear();
                for (int row = 2; row <= lastrow; row++)
                {
                    var name = worksheet.Cells[row, 1].Text;
                    var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                    var up = float.Parse(worksheet.Cells[row, 3].Text);
                    Product.ProductStatus status = ProdStatus(name, filename);
                    Product product = new Product(name, quantity, up, status);
                    products.Add(product);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            // Save to file
            package.SaveAs(new FileInfo(filename));
        }
        
        return products;
    }
    
    //Displays Catalog in Console output.
    public void ShowCatalog(List<Product> products)
    {
        Console.WriteLine("Item    ||    Quantity  ||    Unit Price   ||    Status    ");
        foreach (var product in products)
        {
            Console.WriteLine($"\n{product.ProductName}     ||  {product.Quantity}    ||      {product.UnitPrice}       ||      {ProdStatus(product.ProductName, path)}");
        }
    }

    //indexes product on the catalog list variable.
    // public int ProdIndex(Product product)
    // {
    //     int index = -1;
    //     foreach (var p in products)
    //     {
    //         if (p.ProductName == product.ProductName)
    //         {
    //             index = products.IndexOf(p);
    //             break;
    //         }
    //     }
    //     return index;
    // }
    
    
    //Registering a new product
    public void RegisterProduct(Product item, string filename)
    {
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filename)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                if (CheckAvailibility(item.ProductName, item.Quantity, filename) == (true, true) ||
                    CheckAvailibility(item.ProductName, item.Quantity, filename) == (true, false))
                {
                    Console.WriteLine("\nThe Product is already registered");
                }
                else
                {
                    worksheet.Cells[lastrow+1, 1].Value = item.ProductName;
                    worksheet.Cells[lastrow+1, 2].Value = 0;
                    worksheet.Cells[lastrow+1, 3].Value = item.UnitPrice;
                    
                    Console.WriteLine("The Product was successfully registered");
                }
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            
            // Save to file
            package.SaveAs(new FileInfo(filename));
            //reloading products list to the updated catalog.
            ReadCatalog(filename);
        }
         
    }
    
    
    //order items (add to warehouse) && (log the order hash) --> returns orderHash
    public void Add(List<Product> order, string catalog, string orderlog)
{
    string orderHash = "";
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    // if all products are found
    bool allFound = true;

    using (var package = new ExcelPackage(new FileInfo(catalog)))
    {
        var worksheet = package.Workbook.Worksheets[0];

        //indexes --> last used row
        int lastrow = worksheet.Dimension.End.Row;

        try
        {
            //First, check if all products are found
            foreach (var p in order)
            {
                bool found = false;
                for (int row = 2; row <= lastrow; row++)
                {
                    var name = worksheet.Cells[row, 1].Text;
                    if (name == p.ProductName)
                    {
                        found = true;
                        break;  //Product found
                    }
                }

                if (!found)
                {
                    //If any product is not found, set allFound to false --> exit the loop
                    allFound = false;
                    Console.WriteLine($"Product {p.ProductName} not found in the catalog.");
                    break;  //Stop processing further
                }
            }

            // If all products are found, update catalog
            if (allFound)
            {
                foreach (var p in order)
                {
                    for (int row = 2; row <= lastrow; row++)
                    {
                        var name = worksheet.Cells[row, 1].Text;
                        if (name == p.ProductName)
                        {
                            var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                            worksheet.Cells[row, 2].Value = quantity + p.Quantity;
                            Console.WriteLine($"{p.ProductName} : {p.Quantity} was successfully ordered");
                            break;
                        }
                    }
                }

                // Generate the hash only if the order is valid
                hash.OrderHash(order, catalog, orderlog);

                // Save to file after processing the order
                package.SaveAs(new FileInfo(catalog));

                // Reload the catalog after saving
                ReadCatalog(catalog);
            }
            else
            {
                // If any product was not found, the order does not go through
                Console.WriteLine("Order was not processed due to missing product(s)");
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }
}
    
    
    //ship items (subtract from warehouse) && (log the shipment hash) --> return shipmentHash
    public void Subtract(List<Product> shipment, string catalog, string shipmentlog)
    {
        string shipmentHash = "";
        bool allFound = true;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //making changes to catalog 
        using (var package = new ExcelPackage(new FileInfo(catalog)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                foreach (var o in shipment)
                {
                    bool found = false;
                    for (int row = 2; row <= lastrow; row++)
                    {
                        var name = worksheet.Cells[row, 1].Text;
                        var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                        if (o.ProductName == name)
                        {
                            if (CheckAvailibility(name, o.Quantity, catalog) == (true, true))
                            {
                                found = true; //
                                break;
                            }

                            if (CheckAvailibility(name, o.Quantity, catalog) == (true, false))
                            {
                                found = false;
                                Console.WriteLine($"\nWe dont have enough inventory to Ship: {name}"); 
                                break;
                            }
                        }
                    }
                    
                    if (!found)
                    {
                        allFound = false;
                        Console.WriteLine($"\nThe item:{o.ProductName} was either not found or isn't sufficiently stocked!");
                    }
                }
                
                //if all the products in the list are shippable
                if (allFound)
                {
                    foreach (var s in shipment)
                    {
                        for (int row = 2; row <= lastrow; row++)
                        {
                            var name = worksheet.Cells[row, 1].Text;
                            var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                            if (s.ProductName == name)
                            {
                                worksheet.Cells[row, 2].Value = quantity - s.Quantity;
                                Console.WriteLine(
                                    $"{s.ProductName} : {s.Quantity} was shipped\nCurrent Quantity: {worksheet.Cells[row, 2].Text}");
                            }
                        }
                    }
                    //creating the hash onyl if the shipment is possible
                    hash.ShipmentHash(shipment, catalog, shipmentlog);
                    
                    // Save to file
                    package.SaveAs(new FileInfo(catalog));
                    
                    //reloading products list from catalog.
                    ReadCatalog(catalog);
                }
                //if allfound ==false
                else
                {
                    Console.WriteLine("Shipment was not processed, due to insufficient stock or missing product(s)");
                }
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }

    // (found, IsThereEnoughtoShip)
    public (bool,bool) CheckAvailibility(string itemname, int qty, string filename) 
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filename)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            bool found = false;
            bool availibility = false;
            try
            {
                for (int row = 2; row <= lastrow; row++)
                {
                    var name = worksheet.Cells[row, 1].Text;
                    var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                    
                    found = name == itemname;
                    if (found)
                    {
                        if (quantity - qty >= 0)
                        {
                            availibility = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            // Save to file
            package.SaveAs(new FileInfo(filename));
            return (found, availibility);
        }
    }
    
    //productstatus
    public Product.ProductStatus ProdStatus(string name, string filename)
    {
        var status = new Product.ProductStatus();
        using (var package = new ExcelPackage(new FileInfo(filename)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            bool found = false;
            bool availibility = false;
            try
            {
                for (int row = 2; row <= lastrow; row++)
                {
                    var itemname = worksheet.Cells[row, 1].Text;
                    var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                    
                    found = name == itemname;
                    if (found)
                    {
                        if (quantity == 0)
                        {
                            status = Product.ProductStatus.OutOfStock;
                            break;
                        }

                        if (quantity > 0 && quantity <= 10)
                        {
                            status = Product.ProductStatus.RefillNeeded;
                        }
                        else
                        {
                            status = Product.ProductStatus.InStock;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            // Save to file
            package.SaveAs(new FileInfo(filename));
            
            return status;
        }
    }
    
    
}