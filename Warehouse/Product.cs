namespace Warehouse;

public class Product
{
    public string ProductName;
    public int Quantity;
    public float UnitPrice;
    public enum ProductStatus
    {
        InStock,
        OutOfStock,
        RefillNeeded
    };

    //constructors
    public Product(string productName, int quantity, float unitPrice, ProductStatus status)
    {
        ProductName = productName;
        Quantity = quantity;
        UnitPrice = unitPrice;
    }
    public Product(string productName, float unitPrice)
    {
        ProductName = productName;
        UnitPrice = unitPrice;
    }
    public Product(string productName, int quantity)
    {
        ProductName = productName;
        Quantity = quantity;
    }
    public Product(string productName)
    {
        ProductName = productName;
    }
    public Product(){}
    
    //just reading input for Main Menu
    public Product ReadInput1()
    {
        Console.WriteLine("Register a New Item:\n");
        Console.WriteLine("Item Name:\n");
        string name = Console.ReadLine();
        Console.WriteLine("Item Price:\n");
        float price = float.Parse(Console.ReadLine());
                    
        Product p = new Product(name, price);
        return p;
    }

    public List<Product> ReadInput2()
    {
        List<Product> order = new List<Product>();
        bool done = false;
        string input;
        while (!done)
        {
            Console.WriteLine("Order Item:\n");
            Console.WriteLine("Item Name:\n");
            string name = Console.ReadLine();
            Console.WriteLine("Item Qty:\n");
            int qty = Int32.Parse(Console.ReadLine());
            Console.WriteLine("Type 'done' to finish the order:\n");
            input = Console.ReadLine();
            
            Product p = new Product(name, qty);
            order.Add(p);
            
            if (input == "done")
            {
                done = true;
            }
        }
        return order;
        
    }
    
    public List<Product> ReadInput3()
    {
        List<Product> order = new List<Product>();
        bool done = false;
        string input;
        while (!done)
        {
            Console.WriteLine("Shipment Item:\n");
            Console.WriteLine("Item Name:\n");
            string name = Console.ReadLine();
            Console.WriteLine("Item Qty:\n");
            int qty = Int32.Parse(Console.ReadLine());
            Console.WriteLine("Type 'done' to finish the Shipment:\n");
            input = Console.ReadLine();
            
            Product p = new Product(name, qty);
            order.Add(p);
            
            if (input == "done")
            {
                done = true;
            }
        }
        return order;
        
    }
}