using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using MoreLinq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace ExcelImporter
{
    class StoreExcelToDB
    {

        SqlConnection databaseConnPackages = null;
        SqlConnection databaseConnOrders = null;
        SqlTransaction trans = null;

        List<Supplier> collectionOfSuppliers = null;
        List<string> sprava = new List<string>();

        public StoreExcelToDB()
        {
            string pathPackage = ConfigurationManager.ConnectionStrings["SqlPackages"].ConnectionString;
            databaseConnPackages = new SqlConnection(pathPackage);
            databaseConnPackages.Open();

            string pathOrders = ConfigurationManager.ConnectionStrings["SqlOrders"].ConnectionString;
            databaseConnOrders = new SqlConnection(pathOrders);
            databaseConnOrders.Open();
        }

        public List<string> Store(List<ExcelFile> loadedExcel)
        {

            GetSupplierNamesAndQuantity(loadedExcel);


            trans = databaseConnOrders.BeginTransaction();

            try
            {
                GetSupplierId();


                InsertIntoItems(loadedExcel);
                InsertIntoSuppliers(loadedExcel);
                InsertIntoSuppliersItems(loadedExcel);
                InsertIntoPriceCategories();
                InsertIntoItemPrices(loadedExcel);
                InsertIntoPackages();

                trans.Commit();

                sprava.Add("Ulozenie do databazy prebehlo uspesne");

                WriteToLogFile();
            }
            catch (Exception)
            {
                trans.Rollback();
                throw new Exception();
            }

            return sprava;
        }



        //Metoda vytvori kolekciu dovatelov a mnozstva tovarov
        private void GetSupplierNamesAndQuantity(List<ExcelFile> rows)
        {
            var sw = Stopwatch.StartNew();

            List<string> suppliers = new List<string>();
            collectionOfSuppliers = new List<Supplier>();

            //Cyklus na ziskanie zoznamu vsetkych dodavatelov
            foreach (string item in rows.Select(a => a.supplierName).Distinct())
            {
                suppliers.Add(item);
            }

            foreach (var item in suppliers)
            {
                Supplier supp = new Supplier();
                List<int> local = new List<int>();

                foreach (var quantity in rows.Where(a => a.packageQuantity != -1 && a.supplierName == item).Select(a => a.packageQuantity).Distinct())
                {
                    local.Add(quantity);
                }

                supp.supplierName = item;
                supp.quantity = local;

                collectionOfSuppliers.Add(supp);
            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Vytvorenie kolekcie dodavatelov a zoznamu ich tovarov prebehlo uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");
        }


        private void GetSupplierId()
        {
            SqlCommand cmd1 = new SqlCommand("SELECT MAX(SupplierId)from[OrdersDB].[dbo].[Suppliers]", databaseConnOrders);
            cmd1.Transaction = trans;
            int lastId = (int)cmd1.ExecuteScalar() + 1;


            for (int i = 0; i < collectionOfSuppliers.Count; i++)
            {
                collectionOfSuppliers[i].supplierId = lastId;

                lastId++;
            }

        }


        private void WriteToLogFile()
        {

            string path = @"C:\Users\toth\Desktop\Logs";
            string date = DateTime.Now.ToString("yyyy_MM_dd_HH.mm");
            string fileName = string.Format("{0}_ExcelImporter.txt", date);
            string pathCombination = Path.Combine(path, fileName);
            
            try
            {
                using (var myFile = File.Create(pathCombination))
                {
                    TextWriter sw = new StreamWriter(myFile);

                    foreach (var item in sprava)
                    {
                        sw.WriteLine(item);
                    }

                    sw.Close();
                }
            }

            catch (Exception)
            {

                throw;
            }

        }


        //[OrdersDB].[dbo].[Items]
        private void InsertIntoItems(List<ExcelFile> rows)
        {
            var sw = Stopwatch.StartNew();

            foreach (ExcelFile item in rows)
            {
                SqlCommand cmd1 = new SqlCommand("IF NOT EXISTS (SELECT * FROM [OrdersDB].[dbo].[Items] WHERE ItemID = @ItemID) BEGIN INSERT INTO [OrdersDB].[dbo].[Items] (ItemID, Name, MeasurementUnitID, Active) VALUES (@ItemID, @ItemName, 1, 2) END; ", databaseConnOrders);

                cmd1.Parameters.AddWithValue("@ItemID", item.productId);
                cmd1.Parameters.AddWithValue("@ItemName", item.saleDescription);

                cmd1.Transaction = trans;
                cmd1.ExecuteNonQuery();
            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Insert udajov do tabulky [OrdersDB].[dbo].[Items] prebehol uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");
        }


        //[OrdersDB].[dbo].[Suppliers]
        private void InsertIntoSuppliers(List<ExcelFile> rows)
        {

            var sw = Stopwatch.StartNew();

            foreach (var item in collectionOfSuppliers)
            {
                SqlCommand cmd2 = new SqlCommand("IF NOT EXISTS (SELECT * FROM [OrdersDB].[dbo].[Suppliers] WHERE NAME = '@SupplierName') INSERT INTO [OrdersDB].[dbo].[Suppliers] (SupplierId, Name, StatusUpdates) VALUES (@SupplierId, @SupplierName, 2) ", databaseConnOrders);

                cmd2.Parameters.AddWithValue("@SupplierId", item.supplierId);
                cmd2.Parameters.AddWithValue("@SupplierName", item.supplierName);

                cmd2.Transaction = trans;
                cmd2.ExecuteNonQuery();

            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Insert udajov do tabulky [OrdersDB].[dbo].[Suppliers] prebehol uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");
        }

        //[OrdersDB].[dbo].[SupplierItems]
        private void InsertIntoSuppliersItems(List<ExcelFile> rows)
        {
            var sw = Stopwatch.StartNew();

            foreach (var item in rows)
            {

                SqlCommand cmd1 = new SqlCommand("IF NOT EXISTS (SELECT * FROM [OrdersDB].[dbo].[SupplierItems] WHERE SupplierItemID = @SupplierItemId) INSERT INTO [OrdersDB].[dbo].[SupplierItems] (SupplierID, SupplierItemID, ItemID, SupplierItemName, Active) VALUES(@SupplierId, @SupplierItemId, @ItemId, @SupplierItemName, 2)", databaseConnOrders);

                cmd1.Parameters.AddWithValue("@SupplierId", collectionOfSuppliers.Where(a => a.supplierName.Equals(item.supplierName)).Select(b => b.supplierId).First());
                cmd1.Parameters.AddWithValue("@SupplierItemId", item.vendorProductId);
                cmd1.Parameters.AddWithValue("@ItemId", item.productId);
                cmd1.Parameters.AddWithValue("@SupplierItemName", item.productName);

                cmd1.Transaction = trans;
                cmd1.ExecuteNonQuery();

            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Insert udajov do tabulky [OrdersDB].[dbo].[SupplierItems] prebehol uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");
        }


        //[OrdersDB].[dbo].[Packages]
        private void InsertIntoPackages()
        {
            var sw = Stopwatch.StartNew();

            foreach (var item in collectionOfSuppliers)
            {
                List<int> kolekcia = collectionOfSuppliers.Where(a => a.supplierName == item.supplierName).Select(a => a.quantity).First();

                foreach (int amount in kolekcia)
                {
                    SqlCommand cmd1 = new SqlCommand("INSERT INTO [OrdersDB].[dbo].[Packages] VALUES (@SupplierId, @Description, @Amount)", databaseConnOrders);

                    cmd1.Parameters.AddWithValue("@SupplierId", item.supplierId);

                    cmd1.Parameters.AddWithValue("@Description", string.Format("{0} kusov", amount));
                    cmd1.Parameters.AddWithValue("@Amount", amount);

                    cmd1.Transaction = trans;
                    cmd1.ExecuteNonQuery();
                }

            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Insert udajov do tabulky [OrdersDB].[dbo].[Packages] prebehol uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");
        }


        //[OrdersDB].[dbo].[PriceCategories]
        private void InsertIntoPriceCategories()
        {
            var sw = Stopwatch.StartNew();

            foreach (int item in collectionOfSuppliers.Select(a => a.supplierId))
            {
                SqlCommand cmd1 = new SqlCommand("INSERT INTO [OrdersDB].[dbo].[PriceCategories] (SupplierID, Description) VALUES (@SupplierId, 'Bežná cena')", databaseConnOrders);

                cmd1.Parameters.AddWithValue("@SupplierId", item);

                cmd1.Transaction = trans;
                cmd1.ExecuteNonQuery();
            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Insert udajov do tabulky [OrdersDB].[dbo].[PriceCategories] prebehol uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");
        }


        //[OrdersDB].[dbo].[ItemPrices]
        private void InsertIntoItemPrices(List<ExcelFile> rows)
        {
            var sw = Stopwatch.StartNew();


            foreach (int supplierId in collectionOfSuppliers.Select(a => a.supplierId))
            {
                SqlCommand cmd1 = new SqlCommand("SELECT PriceCategoryID FROM [OrdersDB].[dbo].[PriceCategories] WHERE SupplierID = @SupplierId", databaseConnOrders);
                cmd1.Parameters.AddWithValue("@SupplierId", supplierId);

                cmd1.Transaction = trans;
                int priceCategory = (int)cmd1.ExecuteScalar();

                foreach (var item in rows.DistinctBy(a => a.productId))
                {
                    SqlCommand cmd2 = new SqlCommand("IF NOT EXISTS (SELECT ItemID, PriceCategoryID FROM [OrdersDB].[dbo].[ItemPrices] WHERE ItemID = @ItemId AND PriceCategoryID = @PriceCategory) INSERT INTO [OrdersDB].[dbo].[ItemPrices] VALUES (@ItemId, @PriceCategory, @RetailPrice)", databaseConnOrders);

                    cmd2.Parameters.AddWithValue("@ItemId", item.productId);
                    cmd2.Parameters.AddWithValue("@PriceCategory", priceCategory);
                    cmd2.Parameters.AddWithValue("@RetailPrice", item.retailPrice);

                    cmd2.Transaction = trans;
                    cmd2.ExecuteNonQuery();

                }

            }

            sw.Stop();

            double totalTime = sw.Elapsed.TotalMilliseconds;

            sprava.Add("Insert udajov do tabulky [OrdersDB].[dbo].[ItemPrices] prebehol uspesne");
            sprava.Add(string.Format("Celkovy cas {0} ms", totalTime));
            sprava.Add("---------------");

        }

    }
}
