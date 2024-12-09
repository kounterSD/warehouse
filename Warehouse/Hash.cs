using BCrypt;
using OfficeOpenXml;

namespace Warehouse;

public class Hash
{
    public string GetHash(List<Product> order, string type, int Id)
    {
        string orderstring = $"{type}\n";
        foreach (var item in order)
        {
            orderstring += $"{item.ProductName} : {item.Quantity}\n";
        }
        string hash = BCrypt.Net.BCrypt.HashPassword(orderstring);
        
        return hash;
    }
    
    public string OrderHash(List<Product> order, string catalog, string orderlog)
    {
        string orderHash;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        //logging the order in orderlog.xlsx
        using (var package = new ExcelPackage(new FileInfo(orderlog)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                //generating order hash
                orderHash = GetHash(order, "Order", lastrow+1);
                Console.WriteLine($"Order hash: {orderHash}");
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
        return orderHash;
    }
    
    public string ShipmentHash(List<Product> order, string catalog, string shipmentlog)
    {
        string orderHash;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        //logging the order in orderlog.xlsx
        using (var package = new ExcelPackage(new FileInfo(shipmentlog)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            //indexes -->last used row
            int lastrow = worksheet.Dimension.End.Row;
            try
            {
                //generating order hash
                orderHash = GetHash(order, "Shipment", lastrow+1);
                Console.WriteLine($"Shipment hash: {orderHash}");
                worksheet.Cells[lastrow+1, 1].Value = lastrow-1; 
                worksheet.Cells[lastrow+1, 2].Value = orderHash;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            
            // Save to file
            package.SaveAs(new FileInfo(shipmentlog));
        }
        return orderHash;
    }
}