using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImporter
{
    static class ConvertHelper
    {


        public static Int32 TryGetValueInt32(object testedValue, Int32 fixValue)
        {
            if (testedValue == DBNull.Value || !(testedValue is Double) )
                return fixValue;
            else
                return Convert.ToInt32(testedValue);
        }

        public static Double TryGetValueDouble(object testedValue, Double fixValue)
        {
            if (testedValue == DBNull.Value || !(testedValue is Double))
                return fixValue;
            else
                return Convert.ToDouble(testedValue);
        }

        public static String TryGetValueString(object testedValue, String fixValue)
        {
            if (testedValue == DBNull.Value || !(testedValue is String))
                return fixValue;
            else
                return Convert.ToString(testedValue);
        }

      
    }
}
