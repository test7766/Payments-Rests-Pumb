using System;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.IO;
using System.Linq;
using FormatLibrary;

namespace RestPumb
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.Title = "Залишкі Пумб";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            LoadExcelFile();
            Console.ReadKey();

            void LoadExcelFile()
            {
                //Original Fields
                string[] originalHeaders = { "case#", "CASE_S_ID", "CASE_S_ID full", "CASE_CODE", "Account", "branch", "CUST_CODE", "CIF", "INN", "BIRTHDATE", "Name", "ProdCategory",
            "Product","Open date","Expiration date","DPD","Сума виданого кредиту","Currency","Debt","Outstanding","SA_Debt","Penalties","Основний борн (тіло)",
            "Щомісячний платіж","Queue","Agency","Date transfered to agency","Zone","Legal address","Physical_address","Gen_status3","Bank","Last_pmt_dt_2017","Last_rpc",
            "Days w/o rpc","PayHub_link","Місце роботи позичальника","Назва посади","MOBILE","DOP_RPC","FORGIVE_PAYMENT"};
                //Rests
                string[] rests = { "case#", "CASE_S_ID full", "inn", "Currency", "Debt", "Outstanding", "SA_Debt", "Penalties", "Date transfered to agency", "MOBILE", "DOP_RPC", "FORGIVE_PAYMENT" };

                // relative path
                string currentDirectoryGetXlsx = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\r.xlsx";
                string outDirectoryPayment = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\r.csv";

                ClassHelper classHelper = new ClassHelper();
                try
                {

                    ///check Input xlsx
                    if (!File.Exists(currentDirectoryGetXlsx))
                        throw new Exception($"File {currentDirectoryGetXlsx} is not exist!!!");


                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("filter data field is:");
                    Console.ResetColor();
                    Console.WriteLine(string.Join(" ", rests));


                    using (var package = new ExcelPackage(new FileInfo(currentDirectoryGetXlsx)))
                    {

                        var worksheet = package.Workbook.Worksheets[0];
                        if (worksheet == null)
                            throw new Exception("Worksheet is null");

                        if (worksheet.Rows.Count() == 0)
                            throw new Exception("worksheet.Rows is null");


                        if (worksheet.Columns.Count() == 1)
                            throw new Exception("worksheet.Columns is null");


                        // headers
                        var columnHeaders = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns]
                            .Select(cell => cell.Text.Trim())
                            .ToList();



                        if (columnHeaders.Count() != originalHeaders.Count())
                            throw new Exception("Headers mismatch ");



                        if (!classHelper.CompareHeaders(originalHeaders, columnHeaders))
                            throw new Exception("Invalid file");


                        int originColumns = worksheet.Columns.Count();



                        foreach (var columnIndex in classHelper.GetColumnsToRemove(columnHeaders, rests).OrderByDescending(i => i))
                            worksheet.DeleteColumn(columnIndex, 1);



                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;


                        ///Output Csv

                        using (var writer = new StreamWriter(outDirectoryPayment))
                        {
                            for (int row = 1; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= colCount; col++)
                                {
                                    writer.Write($"{worksheet.Cells[row, col].Text}{(col == colCount ? "" : ";")}");
                                }
                                writer.WriteLine();
                            }
                        }

                        Console.Clear();
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("filter data field is:");
                        Console.ResetColor();
                        Console.WriteLine(string.Join(" ", rests));

                        Console.WriteLine("---------------------");

                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("Headers :");
                        Console.ResetColor();
                        Console.WriteLine(string.Join(" ", columnHeaders));

                        Console.WriteLine("---------------------");

                        Console.Write("Count of rows : "); Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine($"{worksheet.Rows.Count()}");
                        Console.ResetColor();
                        Console.Write("Count of Columns : "); Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine($"{originColumns}");
                        Console.ResetColor();

                        Console.WriteLine("---------------------");

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Data is sucsesufull save in {outDirectoryPayment}");
                        Console.ResetColor();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }



    }
}
