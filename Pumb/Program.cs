
using FormatLibrary;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace Pumb
{
    public class Program
    {
    
        static void Main(string[] args)
        {
         
            Console.Title = "Платежі Пумб";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            LoadExcelFile();
            Console.WriteLine("pres key to continue....");
            Console.ReadKey();

            void LoadExcelFile()
            {

                ClassHelper classHelper = new ClassHelper();
                //Payments
                string[] rests = { "CASE_CONTR_NUM", "date_pay", "pay_cvr_wo_cons", "PACK_ASSIGN_DATE", "CUST_AFM", "Рефинансирование", "Договорное списание" };

                string[] originalHeaders = { "date", "CASE_BRANCH_INTCODE", "case_contr_num", "CUST_CIF", "product", "bucket_before_pmt", "date_pay", "pay_month", "agency",
                "pay_total","debt","pay_cvr_wo_cons","bank","delinquent_days_at_start_month","category","PACK_ASSIGN_DATE","CUST_AFM","Рефинансирование",
                "Договорное списание","case_code","Код плательщика"};

                // relative path
                string currentDirectoryGetXlsx = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\p.xlsx";

                string outDirectoryPayment = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\p.csv";

                try
                {

                    ///check Input xlsx
                    if (!File.Exists(currentDirectoryGetXlsx))
                        throw new Exception($"File {currentDirectoryGetXlsx} is not exist!!!");


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
                                    var column = worksheet.Cells[row, col].Text;
                                    if (col == colCount)
                                    {
                                        if (col == 3)
                                        {
                                            string formatNumber = classHelper.GetFormattedCellValueNumber(worksheet.Cells[row, col].Text);
                                            writer.Write($"{formatNumber};");
                                            Console.Write($"{formatNumber};");
                                        }
                                        else if (col == 2 || col == 4)
                                        {
                                            string newFormatDate = classHelper.FormatDate(worksheet.Cells[row, col].Text);
                                            writer.Write($"{newFormatDate};");
                                            Console.Write($"{newFormatDate};");
                                        }
                                        else
                                        {
                                            writer.Write($"{worksheet.Cells[row, col].Text}");
                                            Console.Write($"{worksheet.Cells[row, col].Text}");
                                        }
                                    }
                                    else
                                    {

                                        if (col == 3)
                                        {
                                            var formatNumber = classHelper.GetFormattedCellValueNumber(worksheet.Cells[row, col].Text);
                                            writer.Write($"{formatNumber};");
                                            Console.Write($"{formatNumber};");
                                        }
                                        else if (col == 2 || col == 4)
                                        {
                                            var newFormatDate = classHelper.GetFormattedCellValueNumber(worksheet.Cells[row, col].Text);
                                            writer.Write($"{newFormatDate};");
                                            Console.Write($"{newFormatDate};");
                                        }
                                        else
                                        {
                                            writer.Write($"{worksheet.Cells[row, col].Text};");
                                            Console.Write($"{worksheet.Cells[row, col].Text};");
                                        }

                                    }
                                }
                                writer.WriteLine();
                                Console.WriteLine();
                            }
                        }
        
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("filter headers:");
                        Console.ResetColor();
                        Console.WriteLine(string.Join(" ", rests));

                        Console.WriteLine("---------------------");

                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("Original headers :");
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
