using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImporter
{
    class ApplicationConfiguration
    {


        public static string[] GetFilledPriceBookColumnNames()
        {

            var defaultFilledColumnNames = new[] {"name", "prod_id", "vend_prod_id", "prod_name", "dept_id", "cat_id", "package_unit_id", "retl_pack_id", "barcode", "sale_desc", "pack_cost", "pack_qty", "retail"};

            string[] columnNames = null;

            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string setting = appSettings["FilledColumnNames"] ?? string.Empty;

                columnNames = setting.Split(new[] { ';' }, StringSplitOptions.None);

                if (columnNames.Length != defaultFilledColumnNames.Length)
                {
                    columnNames = null;
                    throw new Exception("Pocet stlpcov uvedenych v configu nezodpoveda defaultnemu poctu stlpcov");

                }
            }
            catch (ConfigurationErrorsException)
            {
                throw new ConfigurationErrorsException(string.Format("{0} - - - - Nepodarilo sa nacitat vsetky nazvy stlpcov", DateTime.Now.ToLocalTime().ToString()));
            }

            return columnNames ?? defaultFilledColumnNames;
        }
    }
}
