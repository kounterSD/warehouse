using BCrypt;

namespace Warehouse;

public class Hash
{
    public string GetOrderHash(List<Product> order, string type, int Id)
    {
        string orderstring = $"{type}\n";
        foreach (var item in order)
        {
            orderstring += $"{item.ProductName} : {item.Quantity}\n";
        }
        string hash = BCrypt.Net.BCrypt.HashPassword(orderstring);
        
        return hash;
    }
}