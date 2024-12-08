//Hashing orders was not implemented. 

using System;
using System.Data.Common;
using System.Runtime.Intrinsics.Arm;
using OfficeOpenXml;

namespace Warehouse
{
    class Warehouse
    {
        public static void Main()
        {
            //initialising objects.
            Excel excel = new Excel();
            Product product = new Product();
            List<Product> Catalog = new List<Product>();
            List<Product> Order = new List<Product>();
            List<Product> Shipment = new List<Product>();
            string catalogfile = "./../../../warehouse.xlsx";
            string orderfile = "./../../../orderlog.xlsx";
            string shipmentfile = "./../../../shipmentlog.xlsx";
            
            Catalog = excel.ReadCatalog(catalogfile);
            bool exit = false;
            do
            {
                string option =ShowMenu().ToString();
                switch (option)
                {
                    //Show Catalog
                    case "1":
                        excel.ShowCatalog(Catalog);
                        break;
                    
                    //Register a new product for the warehouse
                    case "2":
                        Product p = product.ReadInput1();
                        excel.RegisterProduct(p, catalogfile);
                        break;
                    
                    //Purchase an order --> Items come into the warehouse
                    case "3":
                        Order = product.ReadInput2();
                        Console.WriteLine($"Success!\nOrder Hash: {excel.Add(Order, catalogfile, orderfile)}");
                        break;
                    //Ship an order --> Items go out of the warehouse
                    case "4":
                        Shipment = product.ReadInput3();
                        Console.WriteLine($"Success!\nShipment Hash: {excel.Subtract(Shipment, catalogfile, shipmentfile)}");
                        break;
                    case ("5"):
                        exit = true;
                        break;
                
                }
            } while (!exit);
            
            
        }

        public static int ShowMenu()
        {
            Console.WriteLine("Welcome to the Warehouse!\n");
            Console.WriteLine("1. Show Product Catalog\n2. Register New Item\n3. Purchase an Order(Import)\n4. Ship an Order(Export)\n5. Exit\n");
            int option = Int32.Parse(Console.ReadLine());
            return option;

        }
    }
}