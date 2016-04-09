using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImporter
{
    class ReadExcel
    {
        OleDbConnection excelConn = null;

        public ReadExcel(string filename)
        {
            string connectionString = null;

            if (filename.ToLower().Contains(".xlsx") || filename.ToLower().Contains(".xlsm"))
            {
                connectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=Yes;'",
                            filename);
            }

            else if (filename.ToLower().Contains(".xls"))
            {
                connectionString = string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;'",
            filename);
            }

            excelConn = new OleDbConnection(connectionString);
            excelConn.Open();
        }


        public List<ExcelFile> ExcelReader()
        {
            List<ExcelFile> temp = new List<ExcelFile>();

            string commandString = string.Format("SELECT * FROM [Sheet1$]");
            OleDbCommand command = new OleDbCommand(commandString, excelConn);

            using (OleDbDataReader dataReader = command.ExecuteReader())
            {
                while (dataReader.Read())
                {
                    try
                    {
                        ExcelFile row = this.GetFilledRow(dataReader);
                        if (row.supplierName != null && row.productId != 0 && row.vendorProductId != null && row.productName != null)
                            temp.Add(row);
                    }

                    catch (IndexOutOfRangeException ex)
                    {
                        throw new IndexOutOfRangeException("Zle pomenovany stlpec. Ocakavany nazov stlpca je - - - " + ex.Message);
                    }
                }

            }

            return temp;
        }

        private ExcelFile GetFilledRow(OleDbDataReader dataReader)
        {
            ExcelFile row = new ExcelFile();

            row.supplierName = ConvertHelper.TryGetValueString(dataReader[0], string.Empty).TrimEnd();
            row.productId = ConvertHelper.TryGetValueInt32(dataReader[1], 0);
            row.vendorProductId = ConvertHelper.TryGetValueString(dataReader[2], string.Empty);
            row.productName = ConvertHelper.TryGetValueString(dataReader[3], string.Empty);
            row.departmentId = ConvertHelper.TryGetValueString(dataReader[4], string.Empty);
            row.categoryId = ConvertHelper.TryGetValueString(dataReader[5], string.Empty);
            row.packageUnitId = ConvertHelper.TryGetValueString(dataReader[6], string.Empty);
            row.retailPackageId = ConvertHelper.TryGetValueDouble(dataReader[7], 0.00);
            row.barcode = ConvertHelper.TryGetValueString(dataReader[8], string.Empty);
            row.saleDescription=ConvertHelper.TryGetValueString(dataReader[9], string.Empty);
            row.packageCost = ConvertHelper.TryGetValueDouble(dataReader[10], 0.00);
            row.packageQuantity = ConvertHelper.TryGetValueInt32(dataReader[11], 0);
            row.retailPrice = ConvertHelper.TryGetValueDouble(dataReader[12], 0.00);
            row.vat = findVAT(row.departmentId, row.categoryId);

            return row;
        }

        private int findVAT(string deptId, string catId)
        {
            int vat;

            string code = deptId + "-" + catId;

            List<string> vat10 = new List<string>() { "380-100", "400-400", "410-201", "450-204" };
            List<string> vat0 = new List<string>() { "430-100", "440-201", "440-204", "440-210", "440-213", "440-217", "440-221", "450-201", "460-107", "460-201",
                "460-205", "470-204", "900-100", "900-200", "910-100", "950-100" };

            if (vat10.Contains(code))
            {
                vat = 10;
            }
            else if (vat0.Contains(code))
            {
                vat = 0;
            }
            else
            {
                vat = 20;
            }

            return vat;
        }

    }
}
