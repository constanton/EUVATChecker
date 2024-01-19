using DocumentFormat.OpenXml;

using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Office2016.Excel;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Padi.Vies;


namespace EUVATChecker
{
    class Program {
        static async Task Main(string[] args)
        {
            Console.WriteLine("VATChecker started. Opening the Excel file...");
            //set the filename and location of the Excel spreadsheet
            string fileNamePath = @".\vat.xlsx";

           
            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(fileNamePath, false))
            {
                //initialise counters and lists
                int invalidCount = 0;
                int inactiveCount = 0;

                List<string> invalidList = new List<string>();
                List<string> inactiveList = new List<string>();
                List<string> inactiveVATNames = new List<string>();
                List<string> invalidVATNames = new List<string>();

                //initialise access to Excel data
                Console.WriteLine("\nExcel file found and opened.");
                Console.WriteLine("\nChecking 2nd column values under the first row (assuming header exists in row 1).");
                WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                Console.Write("\nStarting ViesManager......");

                //Instantiate ViesManager from vies-dotnet library
                var viesManager = new ViesManager();
                Console.WriteLine("...Done");

                Console.Write("\nChecking for invalid and inactive VAT numbers......");
                //go through all rows with data
                foreach (Row r in sheetData.Elements<Row>())
                { 
                    //skip the first row
                    if (r.RowIndex.Value != 1) 
                    {
                        //get the 2nd column values (the VAT number) using the custom GetCellValue function
                        string vatnumber = GetCellValue(spreadsheetDoc, r.Descendants<Cell>().ElementAt(1));

                        //get the 3rd column data (the VAT holder name)
                        string vatname = GetCellValue(spreadsheetDoc, r.Descendants<Cell>().ElementAt(2));
                        
                        var validresult = ViesManager.IsValid(vatnumber);

                        //check if the result is valid
                        if(validresult.IsValid == false)
                        {
                            //if the vat number invalid add the VAT number in the list of invalid VAT numbers
                            invalidList.Add(vatnumber);
                            // then add the VAT holder name in the list of invalid VAT number holder names
                            invalidVATNames.Add(vatname);
                            //increment the invalid VAT number counter
                            invalidCount++;
                        }

                        // call the asynchronous function checkActive and return the result
                        var activeresult = await checkActive(viesManager, vatnumber);
                        
                        //check if the VAT number is active
                        if (activeresult == false)
                        {
                            //if the VAT number  is inactive add the VAT number in the list of inactive VAT numbers
                            inactiveList.Add(vatnumber);
                            // then add the VAT holder name in the list of inactive VAT number holder names
                            inactiveVATNames.Add(vatname);
                            //increment the inactive VAT number counter
                            inactiveCount++;
                        }
                    }
                }
                Console.WriteLine("...Done");

                // Show the results
                if (invalidCount == 0)
                {
                    Console.WriteLine("\nAll entries are valid");
                }
                else
                {
                    Console.WriteLine("\nInvalid: " + invalidCount + " found.");
                    Console.WriteLine("List of invalid VAT numbers:");
                    for (int i = 0; i < invalidCount; i++)
                    {
                        Console.WriteLine(invalidList[i] + " " + invalidVATNames[i]);
                    }
                }
                if (inactiveCount == 0)
                {
                    Console.WriteLine("\nAll entries are valid");
                }
                else
                {
                    Console.WriteLine("\nInactive: " + inactiveCount + " found.");
                    Console.WriteLine("List of inactive VAT numbers:");
                    for (int i = 0; i < inactiveCount; i++)
                    {
                        Console.WriteLine(inactiveList[i] + " " + inactiveVATNames[i]);
                    }
                }
            }
            Console.WriteLine(" \r\n \r\n");
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();

        }

        //an asynchronous function to call ViesManager method IsActiveAsync
        async public static Task<bool> checkActive(ViesManager viesmanager, String vatnumber)
        {
            var result = await viesmanager.IsActiveAsync(vatnumber);
            return result.IsValid;
        }

        //a custom function to get the cell value from an Excel spreadsheet cell
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }


    }
}
