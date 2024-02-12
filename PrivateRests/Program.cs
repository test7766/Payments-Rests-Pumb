using System.IO;
using System.Linq;
using System;
using FormatLibrary;
using OfficeOpenXml;

namespace PrivateRests
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.Title = "Залишкі Приват";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            LoadExcelFile();
            Console.WriteLine("pres key to continue....");
            Console.ReadKey();

            void LoadExcelFile()
            {

                ClassHelper classHelper = new ClassHelper();
                //Rests
                string[] restsFiltrHeaders = { "ClientID", "refContract", "Сума прострочки на дату залишків в валюті", "Непрострочений борг, в валюті", "Військовий" };

                string[] originalHeaders = { "Дата формування реєстра", "Дата передачі в аутсорс", "Дата по залишкам", "Каты/Інші кредити ФО і т.д.", "Служба відпрацювання",
                "Назва компанії","ФІО боржника","clientId","refContract","IBAN рахунок для погашення","Карта для погашення","Днів у відпрацюванні","Валюта кредита",
                "ОДБ/П48","Кіл-ть днів прострочки","Дата останнього погашення","Сума всіх надходжень","Загальний борг на дату залишків в грн","Загальний борг на дату залишків в валюті",
                "Сума прострочки на дату залишків в грн","Сума прострочки на дату залишків в валюті","Непрострочений борг, в валюті","непросрочка з мінусом","Борг по тілу на дату залишків в валюті",
                "Борг по процентам на дату залишків в валюті","Борг по комісії/ пеня/ штрафам на дату залишків в валюті","Борг по тілу на дату залишків в грн","Борг по процентам на дату залишків в грн",
                "Борг по комісії/ пеня/ штрафам на дату залишків в грн","Військовий","Реструктуризація доступна для борд/ червоних зон","військові з попереднього"};


                // relative path
                string currentDirectoryGetXlsx = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\r.xlsx";

                string outDirectoryPayment = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\r.csv";

                try
                {
                    Console.WriteLine("Start Proccesing....");
                    ///check Input xlsx
                    if (!File.Exists(currentDirectoryGetXlsx))
                        throw new Exception($"File {currentDirectoryGetXlsx} is not exist!!!");


                    using (var package = new ExcelPackage(new FileInfo(currentDirectoryGetXlsx)))
                    {

                        //check Worksheets counts
                        if (package.Workbook.Worksheets.Count != 2)
                            throw new Exception("Count of Worksheets is mismatch");


                        //check name worksheets must be 2
                        if (package.Workbook.Worksheets[0].Name != "Борг" && package.Workbook.Worksheets[1].Name != "Погашення")
                            throw new Exception("Name of Worksheets isn't correct!! ");


                        //get Worksheets Борг
                        var worksheet = package.Workbook.Worksheets[0];
                        if (worksheet == null)
                            throw new Exception("Worksheet is null");

                        if (worksheet.Rows.Count() == 0)
                            throw new Exception("worksheet.Rows is null");


                        if (worksheet.Columns.Count() == 1)
                            throw new Exception("worksheet.Columns is null");


                        worksheet.DeleteRow(1); //only 2 fields


                        // headers
                        var columnHeaders = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns]
                            .Select(cell => cell.Text.Trim())
                            .ToList();

                        if (columnHeaders.Count() != originalHeaders.Count())
                            throw new Exception("Headers mismatch ");

                        if (!classHelper.CompareHeaders(originalHeaders, columnHeaders))
                            throw new Exception("Invalid file");

                        int OriginalColumnsCount = worksheet.Dimension.Columns;


                        foreach (var columnIndex in classHelper.GetColumnsToRemove(columnHeaders, restsFiltrHeaders).OrderByDescending(i => i))
                            worksheet.DeleteColumn(columnIndex, 1);



                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // Output Csv
                        using (var writer = new StreamWriter(outDirectoryPayment))
                        {
                            for (int row = 1; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= colCount; col++)
                                {
                                    string cellValue = worksheet.Cells[row, col].Text.Trim();

                                    if (col == 3)
                                    {
                                        cellValue = classHelper.GetFormattedCellValueNumber(cellValue);
                                    }
                                    else if (col == 4 && cellValue == "0,00")
                                    {
                                        cellValue = cellValue.Replace("0", string.Empty).Replace(",", string.Empty).Trim();
                                    }

                                    writer.Write($"{cellValue}");

                                    if (col < colCount)
                                    {
                                        writer.Write(";");
                                    }
                                }
                                writer.WriteLine();
                            }
                        }


                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("filter headers:");
                        Console.ResetColor();
                        Console.WriteLine(string.Join(" ", restsFiltrHeaders));

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
