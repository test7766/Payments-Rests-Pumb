using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormatLibrary
{
    public class FormatHelper
    {


        //Format in 3 column 2 700,00 -->2700,00 
        public static string GetFormattedCellValueNumber(string cellValue)
        {
            if (decimal.TryParse(cellValue.Replace(".", ","), out decimal decimalValue))
                return Math.Round(decimalValue, 2).ToString();

            return cellValue;
        }



        public static string FormatDate(string cellValue)
        {
            //create array string

            DateTime newDatetime;
            if (DateTime.TryParseExact(cellValue.Replace('-', '.'), "dd.MM.yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out newDatetime))
            {
                return newDatetime.ToString("dd.MM.yyyy");
            }
            return cellValue.ToString();
        }

    }
}
