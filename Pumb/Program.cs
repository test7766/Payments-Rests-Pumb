
using FormatLibrary;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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
        }

        static void LoadExcelFile()
        {

            //Payments
            string[] rests = { "CASE_CONTR_NUM", "date_pay", "pay_cvr_wo_cons", "PACK_ASSIGN_DATE", "CUST_AFM", "Рефинансирование", "Договорное списание" };


            // relative path
            string currentDirectoryGetXlsx = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\p.xlsx";

            string outDirectoryPayment = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\p.csv";

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

                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("Headers :");
                    Console.ResetColor();
                    Console.WriteLine(string.Join(" ", columnHeaders));
                    List<int> columnsToRemove = new List<int>();

                    for (int index = 1; index <= columnHeaders.Count; index++)
                    {
                        string header = columnHeaders[index - 1];
                        bool shouldRemove = true;
                        ///Filter
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

                    foreach (var columnIndex in columnsToRemove.OrderByDescending(i => i))
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
                                        string formatNumber = FormatHelper.GetFormattedCellValueNumber(worksheet.Cells[row, col].Text);
                                        writer.Write($"{formatNumber};");
                                        Console.Write($"{formatNumber};");
                                    }
                                    else if (col == 2 || col == 4)
                                    {
                                        string newFormatDate = FormatHelper.FormatDate(worksheet.Cells[row, col].Text);
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
                                        var formatNumber = FormatHelper.GetFormattedCellValueNumber(worksheet.Cells[row, col].Text);
                                        writer.Write($"{formatNumber};");
                                        Console.Write($"{formatNumber};");
                                    }
                                    else if (col == 2 || col == 4)
                                    {
                                        var newFormatDate = FormatHelper.GetFormattedCellValueNumber(worksheet.Cells[row, col].Text);
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
