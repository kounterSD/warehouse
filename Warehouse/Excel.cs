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
    public string Add(List<Product> order, string catalog, string orderlog)
    {
        string orderHash;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        //logging the order in orderlog.xlsx
        using (var package = new ExcelPackage(new FileInfo(orderlog)))
        {
            var worksheet = package.Workbook.Worksheets.Add("Orders");
            worksheet.Cells[1, 1].Value = "ID"; 
            worksheet.Cells[1, 2].Value = "Hash";
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                //generating order hash
                orderHash = hash.GetOrderHash(order, "Order", lastrow+1);
                worksheet.Cells[lastrow+1, 1].Value = lastrow-1; 
                worksheet.Cells[lastrow+1, 2].Value = orderHash;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            
            // Save to file
            package.SaveAs(new FileInfo(orderlog));
        }
        
        //making changes to the catalog
        using (var package = new ExcelPackage(new FileInfo(catalog)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                bool found = false;
                for (int row = 2; row <= lastrow; row++)
                {
                    var name = worksheet.Cells[row, 1].Text;
                    var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                    
                    foreach (var p in order)
                    {
                        if (name == p.ProductName)
                        {
                            worksheet.Cells[row, 2].Value = quantity + p.Quantity;
                        }
                    }
                }
                if (found == false)
                {
                    Console.WriteLine("Product not found in the Catalog");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            } 
            // Save to file
            package.SaveAs(new FileInfo(catalog));
            //reloading products list to the updated catalog.
            ReadCatalog(catalog);
        }
        return orderHash;

    } 
    
    //ship items (subtract from warehouse) && (log the shipment hash) --> return shipmentHash
    public string Subtract(List<Product> order, string catalog, string shipmentlog)
    {
        string shipmentHash;
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        //logging the shipment in shipmentlog.xlsx
        using (var package = new ExcelPackage(new FileInfo(shipmentlog)))
        {
            var worksheet = package.Workbook.Worksheets.Add("Shipments");
            worksheet.Cells[1, 1].Value = "ID"; 
            worksheet.Cells[1, 2].Value = "Hash";
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                //generating order hash
                shipmentHash = hash.GetOrderHash(order, "Shipment", lastrow+1);
                worksheet.Cells[lastrow+1, 1].Value = lastrow-1; 
                worksheet.Cells[lastrow+1, 2].Value = shipmentHash;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            
            // Save to file
            package.SaveAs(new FileInfo(shipmentlog));
        }
        
        //making changes to catalog 
        using (var package = new ExcelPackage(new FileInfo(catalog)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                foreach (var o in order)
                {
                    for (int row = 2; row <= lastrow; row++)
                    {
                        var name = worksheet.Cells[row, 1].Text;
                        var quantity = Int32.Parse(worksheet.Cells[row, 2].Text);
                        if (o.ProductName == name)
                        {
                            if (CheckAvailibility(name, o.Quantity, catalog) == (true, true))
                            {
                                worksheet.Cells[row, 2].Value = quantity - o.Quantity;
                                Console.WriteLine($"{o.ProductName} : {o.Quantity} was shipped\nCurrent Quantity: {worksheet.Cells[row, 2].Text}");
                            }

                            if (CheckAvailibility(name, o.Quantity, catalog) == (true, false))
                            {
                                Console.WriteLine("\nWe dont have enough inventory to complete this Shipment");
                            }
                        }
                    }

                    if (CheckAvailibility(o.ProductName, o.Quantity, catalog) == (false, true) ||
                        CheckAvailibility(o.ProductName, o.Quantity, catalog) == (false, false))
                    {
                        Console.WriteLine($"\nThe item:{o.ProductName} was not found in the catalog");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            // Save to file
            package.SaveAs(new FileInfo(catalog));
            //reloading products list to the updated catalog.
            ReadCatalog(catalog);
        }
        
        return shipmentHash;
         
    }
    
    // (found, IsThereEnoughtoOrder)
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