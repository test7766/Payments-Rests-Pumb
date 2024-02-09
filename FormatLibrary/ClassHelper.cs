using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace FormatLibrary
{
    public class ClassHelper
    {

        //Format in 3 column 2 700.00 --> 2700,00 
        public  string GetFormattedCellValueNumber(string cellValue)
        {

            if (decimal.TryParse(cellValue.Replace(".", ","), out decimal decimalValue))
                return decimalValue.ToString("0.00", CultureInfo.InvariantCulture).Replace(".",",");
           
                  
            return cellValue;
        }

        //yyyy-mm-dd
        public  string FormatDate(string cellValue)
        {

            DateTime newDatetime;
            if (DateTime.TryParseExact(cellValue.Replace('-', '.'), "dd.MM.yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out newDatetime))
            {

                var ttt = newDatetime.ToString("dd.MM.yyyy");
                return newDatetime.ToString("dd.MM.yyyy");
            }
            return cellValue.ToString();
        }
        //dd-mm-yyyy --> dd.mm.yyyy 
        public string FormatDateDot(string cellValue)
        {    
            DateTime newDatetime;
            if (DateTime.TryParse(cellValue, out newDatetime))
                return newDatetime.ToString("dd.MM.yyyy");
            
            return cellValue.ToString();
        }

        public  string MobilePhone(string phones) => phones.Replace('|', ',');

        
        public  bool CompareHeaders(string[] originalHeaders, List<string> columnHeaders)
        {
            int j = 0;
            for (int i = 1; i < columnHeaders.Count(); i++)
            {
                for (; j < originalHeaders.Count();)
                    if (string.Equals(columnHeaders[i - 1], originalHeaders[j++], StringComparison.OrdinalIgnoreCase)) break;
                        else return false;               
            }
            return true;
        }


        public  List<int> GetColumnsToRemove(List<string> columnHeaders, string[] rests)
        {
            List<int> columnsToRemove = new List<int>();

            for (int index = 1; index <= columnHeaders.Count; index++)
            {
                string header = columnHeaders[index - 1];
                bool shouldRemove = true;
                foreach (string cutField in rests)
                {
                    if (string.Equals(header, cutField, StringComparison.OrdinalIgnoreCase))
                    {
                        shouldRemove = false;
                        break;
                    }
                }
                if (shouldRemove)
                {
                    columnsToRemove.Add(index);
                }
            }
            return columnsToRemove;
        }



    }
}
