

using System.Collections.Generic;

namespace ExcelImporter
{
    class ExcelFile
    {
        public string supplierName { get; set; }
        public int productId { get; set; }
        public string vendorProductId { get; set; }
        public string productName { get; set; }
        public string departmentId { get; set; }
        public string categoryId { get; set; }
        public string packageUnitId { get; set; }
        public double retailPackageId { get; set; }
        public string barcode { get; set; }
        public string saleDescription { get; set; }
        public double packageCost { get; set; }
        public int packageQuantity { get; set; }
        public double retailPrice { get; set; }
        public int vat { get; set; }
    }

    class Supplier
    {
        public string supplierName { get; set; }
        public List<int> quantity { get; set; }
        public int supplierId { get; set; }
    }

}
