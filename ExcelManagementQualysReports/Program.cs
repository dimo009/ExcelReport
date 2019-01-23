using ExcelManagementQualysReports.Constants;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ExcelManagementQualysReports
{
    class Program
    {
        static void Main(string[] args)
        {
            var clock = new Stopwatch();

            clock.Start();

            FileInfo file = new FileInfo(PathConstants.rawDataPath);

            //it is hardcoded for testing purposes. When the program is ready, use the DateTime.Now.Month-1
            int currentMonth = DateTime.Now.Month;
            

            //int currentMonth = Convert.ToInt32(DateTime.Now.Month-1);
            int currentYear = Convert.ToInt32(DateTime.Now.Year);
            


            using (ExcelPackage package = new ExcelPackage(file))
            {
                var program = new Program();


                FilterESLdataAndCopytToAnotherSheet(package); //1

                FilterDataAndCopyToAnotherSheet(currentMonth, currentYear, package);  //2

                AddAdditionalColumsForTheMatchedData(package); //3

                FillInTheData(package, program); //4

                CreatePivotTable(package);  //5

                clock.Stop();
                Console.WriteLine("Time elapsed: {0:hh\\:mm\\:ss}", clock.Elapsed);

            }
        }

        private static void FillInTheData(ExcelPackage package, Program program)
        {

            program = new Program();

            //Selecting the "FilteredData" sheet
            ExcelWorksheet filteredData = package.Workbook.Worksheets[4];

            //selecting the ESL sheet

            ExcelWorksheet ESL = package.Workbook.Worksheets[3];



            //check how many row and colums there is in the filteredData sheet

            int rowCount = filteredData.Dimension.Rows; //3430
            int colCount = filteredData.Dimension.Columns; //35

            int rowCountESL = ESL.Dimension.Rows; //118957
            int colCountESL = ESL.Dimension.Columns;//10



            //Adding two 0 on the first two indeces because the List starts from 0 and the excel from 1. The first row of the excel are the headers -> so two zeros to compensate that

            var IPvaluesFromESL = new List<string>() { "0", "0" };
            var consoleIPValuesESL = new List<string>() { "0", "0" };

            // Filling in the IP and ConsoleIP values from ESL
            // initial value of i is 2, because in excel the enumeration starts from 1. In the first row are the headers of the excelsheet

            for (int i = 2; i <= rowCountESL - 2; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(ESL.Cells[i, 9].Value)))
                {
                    IPvaluesFromESL.Add(Convert.ToString(ESL.Cells[i, 9].Value));
                }
                else
                {
                    IPvaluesFromESL.Add("0");
                }
                if (!string.IsNullOrEmpty(Convert.ToString(ESL.Cells[i, 10].Value)))
                {
                    consoleIPValuesESL.Add(Convert.ToString(ESL.Cells[i, 10].Value));
                }
                else
                {
                    consoleIPValuesESL.Add("0");
                }

            }

            // Iteration through the IP values from the Filtered data
            for (int row = 2; row <= rowCount; row++)
            {
                string ip = Convert.ToString(filteredData.Cells[row, 1].Value);

                //finding the index of an existing IP
                int index = IPvaluesFromESL.FindIndex(io => io.Equals(ip, StringComparison.Ordinal));

                //finding the index of repeatable IPs
                //var indeces = Enumerable.Range(0, IPvaluesFromESL.Count).Where(i => IPvaluesFromESL[i] == ip).ToList();


                int colIndex = 2;

                if (index == -1)
                {
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    Console.WriteLine($"Raw {row} was filled in with Empty value");
                    continue;
                }

                //there is no need to check the server status anymore because the data from ESL was filtered in one of the previous steps
                //string serverStatus = Convert.ToString(ESL.Cells[index, 5].Value);

                colIndex = 2;

                if (program.CheckIfESLValuesContainTheIP(ip, IPvaluesFromESL))

                {
                    //INterchanged values

                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 1].Value);
                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 4].Value);
                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 5].Value);
                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 3].Value);
                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 7].Value);
                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 6].Value);
                    filteredData.Cells[row, colIndex++].Value = Convert.ToString(ESL.Cells[index, 8].Value);
                    Console.WriteLine($"Raw {row} was filled in with IP value");
                    colIndex = 2;


                }
                else
                {

                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    filteredData.Cells[row, colIndex++].Value = "ZZZZZ";
                    Console.WriteLine($"Raw {row} was filled in with Empty value");
                }


            }

            package.Save();

            for (int row2 = 2; row2 < rowCount; row2++)
            {
                string ip = Convert.ToString(filteredData.Cells[row2, 1].Value);
                int indexFromConsoleIPvalues = consoleIPValuesESL.FindIndex(io => io.Equals(ip, StringComparison.Ordinal));

                // finding all indexes of repeatable ConsoleIPs

                //List<int> indecesConsoleIP = Enumerable.Range(0, consoleIPValuesESL.Count).Where(i => consoleIPValuesESL[i] == ip).ToList();

                if (Convert.ToString(filteredData.Cells[row2, 2].Value) == "ZZZZZ" && indexFromConsoleIPvalues != -1)
                {

                    int colIndex = 2;

                    //INterchanged values

                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 1].Value);
                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 4].Value);
                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 5].Value);
                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 3].Value);
                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 7].Value);
                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 6].Value);
                    filteredData.Cells[row2, colIndex++].Value = Convert.ToString(ESL.Cells[indexFromConsoleIPvalues, 8].Value);
                    Console.WriteLine($"Raw {row2} was filled in with CONSOLE_IP VALUES...");
                    colIndex = 2;



                }
            }


            package.Save();
        }

        private static void CreatePivotTable(ExcelPackage package)
        {
            Console.WriteLine("Generating a pivot table...");
            // XmlDocument xdoc = package.Workbook.Worksheets[4].WorksheetXml;
            //var nsm = new System.Xml.XmlNamespaceManager(xdoc.NameTable);
            ExcelWorksheet pivotSheet = package.Workbook.Worksheets.Add("Pivot");
            ExcelWorksheet dataSourceSheet = package.Workbook.Worksheets[4];


            //define the data range on the source sheet
            var dataRange = dataSourceSheet.Cells[dataSourceSheet.Dimension.Address];

            //create the pivot table
            var pivotTable = pivotSheet.PivotTables.Add(pivotSheet.Cells["B3"], dataRange, "PivotTable");

            //Devining Pivot Table Style
            pivotTable.TableStyle = OfficeOpenXml.Table.TableStyles.Light1;

            //adding filter field
            var modelField = pivotTable.Fields["System name"];
            ExcelPivotTableField filterField = pivotTable.PageFields.Add(modelField);
            modelField.Sort = eSortType.None;


            //Row fields
            ExcelPivotTableField severityField = pivotTable.RowFields.Add(pivotTable.Fields["Severity"]);
            ExcelPivotTableField titleField = pivotTable.RowFields.Add(pivotTable.Fields["Title"]);
            pivotTable.DataOnRows = false;

            //data fields
            var valuesField = pivotTable.DataFields.Add(pivotTable.Fields["System name"]);                     // pivotTable.DataFields.Add(pivotTable.Fields["System name"]);
            valuesField.Name = "#Affected devices";
            valuesField.Function = DataFieldFunctions.Count;
            valuesField.Format = "#,##0_);(#,##0)";

            //----> Sorting
            //ExcelPivotTableField fieldToSort = titleField;

            //var dataField = valuesField;

            //var xdoc = pivotTable.PivotTableXml;

            //var nsm = new XmlNamespaceManager(xdoc.NameTable);

            //bool descending = true;

            //// "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            //var schemaMain = xdoc.DocumentElement.NamespaceURI;
            //if (nsm.HasNamespace("x") == false)
            //    nsm.AddNamespace("x", schemaMain);

            //// <x:pivotField sortType="descending">
            ////Throw error
            //var pivotField = xdoc.SelectSingleNode(
            //    "/x:pivotTableDefinition/x:pivotFields/x:pivotField[position()=" + (fieldToSort.Index + 1) + "]",
            //    nsm
            //);



            //----> END of sorting

            package.Save();
            Console.WriteLine("Pivot table has been created.");
        }

        private static void FilterESLdataAndCopytToAnotherSheet(ExcelPackage package)
        {
            ExcelWorksheet ESLworksheet = package.Workbook.Worksheets[2];
            ExcelWorksheet filteredESL = package.Workbook.Worksheets.Add("FilteredESL");
            int rowCount = ESLworksheet.Dimension.Rows;
            int colCount = ESLworksheet.Dimension.Columns;
            int startingIndexOfTheFilteredESL = 2;
            var permittedValuesOfTheServers = new List<string> { "move to production", "in production", "installed in DC", "delivered", "ordered" };

            //Removed the first for, no need to be here
            //for (int i = 1; i <= 1; i++)
            //{
            for (int col = 1; col <= colCount; col++)
            {
                filteredESL.Cells[1, col].Value = ESLworksheet.Cells[1, col].Value;
            }
            //}

            string serverStatus;

            Console.WriteLine($"Reading data from ESL and filtering ... Please wait!");

            for (int row = 2; row <= rowCount; row++)
            {

                serverStatus = Convert.ToString(ESLworksheet.Cells[row, 4].Value);

                if (permittedValuesOfTheServers.Any(s => s.Equals(serverStatus)))
                {
                    //removing the second for cycle because it is not needed
                    //for (int i = row; i <= row; i++)
                    //{
                    for (int col = 1; col <= 10; col++)
                    {
                        filteredESL.Cells[startingIndexOfTheFilteredESL, col].Value = ESLworksheet.Cells[row, col].Value;

                    }
                    startingIndexOfTheFilteredESL++;
                    ////}
                }
            }


            package.Save();
            Console.WriteLine("The data has been saved!");
        }

        private bool CheckIfESLValuesContainTheIP(string ip, List<string> IpvaluesFromESL)
        {
            if (!IpvaluesFromESL.Contains(ip))
            {
                return false;
            }
            return true;
        }

        private static void AddAdditionalColumsForTheMatchedData(ExcelPackage package)
        {
            //selecting the filtered data
            ExcelWorksheet worksheet = package.Workbook.Worksheets[4];

            //Add the additional columns

            string[] namesForColums = new string[] { "System name", "System status", "System type", "OS Version", "Technical owner", "Downtime contact", "IP Type" };

            worksheet.InsertColumn(2, 7);

            int colIndex = 2; // the starting position of the new columns


            //range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            //range.Style.Font.Color.SetColor(Color.White);


            Console.WriteLine("Adding additional colums for the data matched from ESL... Please wait!");
            for (int i = 0; i < namesForColums.Length; i++)
            {
                worksheet.Cells[1, colIndex].Value = namesForColums[i];
                worksheet.Cells[1, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, colIndex].Style.Fill.BackgroundColor.SetColor(Color.Black);
                worksheet.Cells[1, colIndex].Style.Font.Color.SetColor(Color.White);
                colIndex++;

            }

            int startRowIndex = 1;
            int startColIndex = 1;
            int totalColumnsCount = worksheet.Dimension.Columns;

            using (ExcelRange autoFilterCells = worksheet.Cells[startRowIndex, startColIndex, startRowIndex, totalColumnsCount])
            {
                autoFilterCells.AutoFilter = true;
            }


            package.Save();

            Console.WriteLine("Columns were adddded. The new format is being saved...");
        }

        private static void FilterDataAndCopyToAnotherSheet(int currentMonth, int currentYear, ExcelPackage package)
        {
            ExcelWorksheet rawDataSheet = package.Workbook.Worksheets[1];
            //Adding new sheet

            ExcelWorksheet fiteredDataSheet = package.Workbook.Worksheets.Add("FilteredData");

            //counting the number of rows and colums of the RawData sheet
            int rowCount = rawDataSheet.Dimension.Rows;
            int colCount = rawDataSheet.Dimension.Columns;
            //to start copying the data from the second row in the FilteredData Worksheet
            int startingRowIndex = 2;

            //removing the first for cycle because it is not needed. This is used to set the headers of the new worksheet
            for (int col = 1; col <= 28; col++)
            {
                fiteredDataSheet.Cells[1, col].Value = rawDataSheet.Cells[1, col].Value;
            }

            string severity;

            Console.WriteLine($"Reading data and filtering ... Please wait!");

            for (int row = 2; row <= rowCount; row++)
            {

                severity = Convert.ToString(rawDataSheet.Cells[row, 12].Value);
                string[] dateNew = Convert.ToString(rawDataSheet.Cells[row, 18].Value).Split(new char[] { '/', ' ' }, StringSplitOptions.RemoveEmptyEntries);

                // DateTime fileDate = DateTime.Parse(dateNew);
                //var date = double.Parse(Convert.ToString(worksheet.Cells[row, 18].Value));
                //lastDetected = DateTime.FromOADate(dateNew);//Convert.ToDateTime(date); //date, "M-dd-YYYY", CultureInfo.InvariantCulture);
                var month = int.Parse(dateNew[0]);   //Convert.ToDateTime(date); //date, "M-dd-YYYY", CultureInfo.InvariantCulture);
                var year = int.Parse(dateNew[2]);

                if ((severity == "5" || severity == "4") & (month == currentMonth && year == currentYear))
                {

                    //removing the first for cycle because it is not needed. The row number is taken from the for cycle above

                    //Console.WriteLine($"Copying row {row} ...");
                    for (int col = 1; col <= 28; col++)
                    {
                        fiteredDataSheet.Cells[startingRowIndex, col].Value = rawDataSheet.Cells[row, col].Value;
                    }
                    startingRowIndex++;
                }
            }

            package.Save();

            Console.WriteLine("The filtered raw data has been saved!");
        }
    }
}


