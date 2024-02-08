using System.IO;
using System.Linq;
using System;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using FormatLibrary;

namespace PrivatPayments
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Платежі Приват";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            LoadExcelFile();
            Console.WriteLine("pres key to continue....");
            Console.ReadKey();

            void LoadExcelFile()
            {

                ClassHelper classHelper = new ClassHelper();
                //Payments
                string[] paymentsFiltrHeaders = { "ClientID", "Refcontract", "Дата проводки", "Сума реального погашення прострочки в грн", "Автоматичне погашення з рахунків боржника" };

                string[] originalHeaders = { "Дата по залишкам", "Служба відпрацювання", "Назва компанії", "ФІО боржника", "ClientID", "Refcontract", "Валюта кредита",
                    "Дата проводки","Просрочка по активу на початок дати проводки в валюті кредита","Борг по активу на початок дати проводки в валюті кредита",
                "Просрочка по активу на початок дати проводки в грн","Сума платежу","Сума реального погашення прострочки в грн","Курс валюти","Кіл-ть днів прострочки на дату надходження",
                "Автоматичне погашення з рахунків боржника"};

                string replaceLastField = "Автоматичне погашення з рахунків боржника";
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

                        ExcelWorksheets countSheets = package.Workbook.Worksheets;

                        //check Worksheets counts
                        if (package.Workbook.Worksheets.Count !=2)
                            throw new Exception("Count of Worksheets is mismatch");



                        //check name worksheets must be 2
                        if (package.Workbook.Worksheets[0].Name != "Борг" && package.Workbook.Worksheets[1].Name != "Погашення")
                            throw new Exception("Name of Worksheets isn't correct!! ");


                        //get Worksheets Погашення
                        var worksheet = package.Workbook.Worksheets[1];
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

                        int OriginalColumnsCount = worksheet.Dimension.Columns;
                      

                        foreach (var columnIndex in classHelper.GetColumnsToRemove(columnHeaders, paymentsFiltrHeaders).OrderByDescending(i => i))
                            worksheet.DeleteColumn(columnIndex, 1);



                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;


                        var aaa = worksheet.Rows.Count();

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
                                        if (col==3)
                                        {
                                            var resultFormatDotDate = classHelper.FormatDateDot(worksheet.Cells[row, col].Text);
                                            writer.Write($"{resultFormatDotDate}");
                                            Console.Write($"{resultFormatDotDate}");
                                        }
                                        else if (col == 5)
                                        {
                                            var resultReplacementDate = worksheet.Cells[row, col].Text.Replace("0", string.Empty);
                                            writer.Write($"{resultReplacementDate}");
                                            Console.Write($"{resultReplacementDate}");
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
                                            var resultFormatDotDate = classHelper.FormatDateDot(worksheet.Cells[row, col].Text);
                                            writer.Write($"{resultFormatDotDate};");
                                            Console.Write($"{resultFormatDotDate};");
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
                        Console.WriteLine(string.Join(" ", paymentsFiltrHeaders));

                        Console.WriteLine("---------------------");

                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("Original headers :");
                        Console.ResetColor();
                        Console.WriteLine(string.Join(" ", columnHeaders));

                        Console.WriteLine("---------------------");

                        Console.Write("Count of rows : "); Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine($"{rowCount}");
                        Console.ResetColor();
                        Console.Write("Count of Columns : "); Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine($"{OriginalColumnsCount}");
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
