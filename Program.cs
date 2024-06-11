using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using ExcelDataReader;
using Microsoft.SharePoint.Client;
using OfficeOpenXml;
using File = System.IO.File;

namespace TravelDates
{
    class Program
    {
        static void Main(string[] args)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string siteUrlHR = "https://sonyapc.sharepoint.com/sites/S039-037-ConnectTest";
            string username = "Connectadmin@sony.onmicrosoft.com";
            string password = "THX@v0lum3";
            string csvFilePath = @"C:\Users\7000036422\Documents\TravelDates\connect_travel_view_2024061012000000200000000.csv";//connect_travel_view_2024052005000000100000000.csv";// filePath
            string filePath = @"C:\Users\7000036422\Documents\TravelDates\connect_travel_view_2024061012000000200000000.xlsx";

            string emp_id = "";
            string value8 = "";


            ConvertCsvToXlsx(csvFilePath, filePath);

            // Create dictionary to store counts of dates
            Dictionary<DateTime, int> dateCounts = new Dictionary<DateTime, int>();
            DataTable dummyDataTable = new DataTable();
            DataTable dummypnaposTable = new DataTable();
            // Create a dummy DataTable

            //DataTable dummyDataTable = new DataTable();
            dummyDataTable.Columns.Add("Country", typeof(string));
            dummyDataTable.Columns.Add("Month", typeof(string));
            dummyDataTable.Columns.Add("Company", typeof(string));
            string[] numberWords = { "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty", "Twentyone", "Twentytwo", "Twentythree", "Twentyfour", "Twentyfive", "Twentysix", "Twentyseven", "Twentyeight", "Twentynine", "Thirty", "Thirtyone" };

            foreach (string word in numberWords)
            {
                dummyDataTable.Columns.Add(word, typeof(int)); // Add columns to the DataTable for number words
            }


            dummyDataTable.Columns.Add("MTD", typeof(int));
            dummyDataTable.Columns.Add("FromDate", typeof(string));
            dummyDataTable.Columns.Add("ToDate", typeof(string));


            //pnapostable
            dummypnaposTable.Columns.Add("Country", typeof(string));
            dummypnaposTable.Columns.Add("Month", typeof(string));
            dummypnaposTable.Columns.Add("Company", typeof(string));
            foreach (string word in numberWords)
            {
                dummypnaposTable.Columns.Add(word, typeof(int)); // Add columns to the DataTable for number words
            }


            dummypnaposTable.Columns.Add("MTD", typeof(int));
            dummypnaposTable.Columns.Add("FromDate", typeof(string));
            dummypnaposTable.Columns.Add("ToDate", typeof(string));



            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    DataTable table = result.Tables[0];
                    DataTable filteredSETable = null;
                    DataTable pnaTable = null;

                    // Apply filter to select specific rows
                    var filteredSE = table.AsEnumerable()
                      .Where(row =>
                      {
                          string auth_itemcat_id = row.Field<string>("auth_itemcat_id")?.Trim();
                          string end_status = row.Field<string>("end_status")?.Trim();
                          value8 = row.Field<string>("value8")?.Trim();
                          emp_id = row.Field<string>("emp_id")?.Trim();
                          // If value8 contains "ID"
                          return auth_itemcat_id != null &&
                                 end_status.StartsWith("Approved", StringComparison.OrdinalIgnoreCase) &&
                                 (auth_itemcat_id.StartsWith("TAC1", StringComparison.OrdinalIgnoreCase) ||
                                  auth_itemcat_id.StartsWith("TAC1A", StringComparison.OrdinalIgnoreCase));

                      })
                      .ToList();

                    if (filteredSE.Any())
                    {
                        filteredSETable = filteredSE.CopyToDataTable();
                        Console.WriteLine(filteredSE);
                    }

                    // Create a list to store dates and their corresponding values
                    List<(DateTime Date, int Value)> dateList = new List<(DateTime, int)>();

                    // Iterate through filtered rows
                    foreach (var row in filteredSE)
                    {
                        DateTime fromDateRow, toDateRow;

                        // Parse "from_date" field
                        if (!DateTime.TryParseExact(row.Field<string>("from_date"), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fromDateRow))
                        {
                            // Handle parsing error for "from_date"
                            Console.WriteLine($"Error parsing 'from_date' for row: {row}");
                        }

                        // Parse "to_date" field
                        if (!DateTime.TryParseExact(row.Field<string>("to_date"), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out toDateRow))
                        {
                            // Handle parsing error for "to_date"
                            Console.WriteLine($"Error parsing 'to_date' for row: {row}");
                        }



                        string value81 = row.Field<string>("value8")?.Trim();

                        Dictionary<string, (DateTime FromDate, List<DateTime> Dates)> datesByMonth = new Dictionary<string, (DateTime, List<DateTime>)>();

                        // Iterate through each month within the range of the current row
                        for (DateTime date = fromDateRow.Date; date <= toDateRow.Date; date = date.AddMonths(1).Date)
                        {
                            // Determine the start and end dates of the month within the specified range
                            DateTime startDateOfMonth = new DateTime(date.Year, date.Month, 1);
                            DateTime endDateOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

                            // Adjust start and end dates based on from_date and to_date
                            if (date == fromDateRow.Date)
                            {
                                startDateOfMonth = fromDateRow;
                            }
                            if (date.AddMonths(1).Date > toDateRow.Date)
                            {
                                endDateOfMonth = toDateRow;
                            }

                            // Add each date within the month to the list for the corresponding month
                            for (DateTime d = startDateOfMonth; d <= endDateOfMonth; d = d.AddDays(1))
                            {
                                // Get the month key (e.g., "January-2020")
                                string monthKey = d.ToString("MMMM-yyyy");

                                // Add the date to the list for the corresponding month
                                if (!datesByMonth.ContainsKey(monthKey))
                                {
                                    datesByMonth[monthKey] = (fromDateRow, new List<DateTime>());
                                }
                                datesByMonth[monthKey].Dates.Add(d);
                            }
                        }




                        // Populate the dummy DataTable
                        foreach (var monthEntry in datesByMonth)
                        {
                            // Add a new row for each month
                            DataRow newRow = dummyDataTable.NewRow();
                            if (value81.StartsWith("CN00ID"))
                            {
                                newRow["Country"] = "Indonesia";
                            }
                            else if (value81.StartsWith("CN00MY"))
                            {
                                newRow["Country"] = "Malaysia";
                            }
                            else if (value81.StartsWith("CN00VN"))
                            {
                                newRow["Country"] = "Vietnam";
                            }
                            else if (value81.StartsWith("CN00PH"))
                            {
                                newRow["Country"] = "Philippines";
                            }
                            else if (value81.StartsWith("CN00TH"))
                            {
                                newRow["Country"] = "Thailand";
                            }
                            else if (value81.StartsWith("CN00IN"))
                            {
                                newRow["Country"] = "India";
                            }



                            newRow["Month"] = monthEntry.Key;
                            if (emp_id.Contains("SS") || emp_id.Contains("SG") || emp_id.Contains("SI") || emp_id.Contains("SJ") || emp_id.Contains("SH") || emp_id.Contains("SL") || emp_id.Contains("ST") || emp_id.Contains("SU") || emp_id.Contains("SV") || emp_id.Contains("SX") || emp_id.Contains("SN"))
                            {
                                newRow["Company"] = "SES";
                            }
                            else if (emp_id.Contains("SE") || emp_id.Contains("SM"))
                            {
                                newRow["Company"] = "SEAP";
                            }

                            for (int i = 1; i <= 31; i++)
                            {
                                newRow[numberWords[i - 1]] = 0; // Use word equivalents instead of numerical values
                            }
                            foreach (DateTime date in monthEntry.Value.Dates)
                            {
                                // Convert the day number to its word equivalent and set the value to 1
                                string dayWord = numberWords[date.Day - 1]; // -1 because arrays are zero-based
                                newRow[dayWord] = 1;
                            }


                            newRow["MTD"] = monthEntry.Value.Dates.Count;
                            newRow["FromDate"] = fromDateRow.ToString("dd-MM-yyyy");
                            newRow["ToDate"] = toDateRow.ToString("dd-MM-yyyy");

                            // Add the row to the DataTable
                            dummyDataTable.Rows.Add(newRow);
                        }

                    }





                    //pna

                    var filteredpna = table.AsEnumerable()
                      .Where(row =>
                      {
                          string auth_itemcat_id = row.Field<string>("auth_itemcat_id")?.Trim();
                          string end_status = row.Field<string>("end_status")?.Trim();
                          value8 = row.Field<string>("value8")?.Trim();
                          emp_id = row.Field<string>("emp_id")?.Trim();
                          
                          // If value8 contains "ID"
                          return auth_itemcat_id != null &&
                                  value8.StartsWith("CN00ID", StringComparison.OrdinalIgnoreCase) &&
                                 end_status.StartsWith("Approved", StringComparison.OrdinalIgnoreCase) &&
                                 (auth_itemcat_id.StartsWith("TAC1", StringComparison.OrdinalIgnoreCase) ||
                                  auth_itemcat_id.StartsWith("TAC1A", StringComparison.OrdinalIgnoreCase));
                                  

                      })
                      .ToList();

                    if (filteredpna.Any())
                    {
                        pnaTable = filteredpna.CopyToDataTable();
                        Console.WriteLine(filteredpna);
                        // Console.ReadKey();

                    }

                    // Create a list to store dates and their corresponding values
                    List<(DateTime Date, int Value)> dateList1 = new List<(DateTime, int)>();

                    // Iterate through filtered rows
                    foreach (var row in filteredpna)
                    {
                        DateTime fromDateRow, toDateRow;

                        // Parse "from_date" field
                        if (!DateTime.TryParseExact(row.Field<string>("from_date"), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fromDateRow))
                        {
                            // Handle parsing error for "from_date"
                            Console.WriteLine($"Error parsing 'from_date' for row: {row}");
                        }

                        // Parse "to_date" field
                        if (!DateTime.TryParseExact(row.Field<string>("to_date"), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out toDateRow))
                        {
                            // Handle parsing error for "to_date"
                            Console.WriteLine($"Error parsing 'to_date' for row: {row}");
                        }



                        string value81 = row.Field<string>("value8")?.Trim();
                        string value23 = row.Field<string>("value23")?.Trim();
                        string value24 = row.Field<string>("value24")?.Trim();
                        string value25 = row.Field<string>("value25")?.Trim();
                        string value26 = row.Field<string>("value26")?.Trim();
                        string value21 = row.Field<string>("value21")?.Trim();
                        string value22 = row.Field<string>("value22")?.Trim();
                        string value29 = row.Field<string>("value29")?.Trim();
                        string value30 = row.Field<string>("value30")?.Trim();

                        Dictionary<string, (DateTime FromDate, List<DateTime> Dates)> datesByMonth = new Dictionary<string, (DateTime, List<DateTime>)>();

                        // Iterate through each month within the range of the current row
                        for (DateTime date = fromDateRow.Date; date <= toDateRow.Date; date = date.AddMonths(1).Date)
                        {
                            // Determine the start and end dates of the month within the specified range
                            DateTime startDateOfMonth = new DateTime(date.Year, date.Month, 1);
                            DateTime endDateOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

                            // Adjust start and end dates based on from_date and to_date
                            if (date == fromDateRow.Date)
                            {
                                startDateOfMonth = fromDateRow;
                            }
                            if (date.AddMonths(1).Date > toDateRow.Date)
                            {
                                endDateOfMonth = toDateRow;
                            }

                            // Add each date within the month to the list for the corresponding month
                            for (DateTime d = startDateOfMonth; d <= endDateOfMonth; d = d.AddDays(1))
                            {
                                // Get the month key (e.g., "January-2020")
                                string monthKey = d.ToString("MMMM-yyyy");

                                // Add the date to the list for the corresponding month
                                if (!datesByMonth.ContainsKey(monthKey))
                                {
                                    datesByMonth[monthKey] = (fromDateRow, new List<DateTime>());
                                }
                                datesByMonth[monthKey].Dates.Add(d);
                            }
                        }




                        // Populate the dummy DataTable
                        foreach (var monthEntry in datesByMonth)
                        {
                            // Add a new row for each month
                            DataRow newRow = dummypnaposTable.NewRow();
                            if (value23.StartsWith("Y")|| value24.StartsWith("Y")|| value25.StartsWith("Y")|| value26.StartsWith("Y"))
                            {
                                newRow["Country"] = "PnA";
                            }
                            if(value21.StartsWith("Y") || value22.StartsWith("Y") || value29.StartsWith("Y") || value30.StartsWith("Y"))
                            {
                                newRow["Country"] = "POS";
                            }
                     

                            newRow["Month"] = monthEntry.Key;
                            if (emp_id.Contains("SS") || emp_id.Contains("SG") || emp_id.Contains("SI") || emp_id.Contains("SJ") || emp_id.Contains("SH") || emp_id.Contains("SL") || emp_id.Contains("ST") || emp_id.Contains("SU") || emp_id.Contains("SV") || emp_id.Contains("SX") || emp_id.Contains("SN"))
                            {
                                newRow["Company"] = "SES";
                            }
                            if (emp_id.Contains("SE") || emp_id.Contains("SM"))
                            {
                                newRow["Company"] = "SEAP";
                            }

                            for (int i = 1; i <= 31; i++)
                            {
                                newRow[numberWords[i - 1]] = 0; // Use word equivalents instead of numerical values
                            }
                            foreach (DateTime date in monthEntry.Value.Dates)
                            {
                                // Convert the day number to its word equivalent and set the value to 1
                                string dayWord = numberWords[date.Day - 1]; // -1 because arrays are zero-based
                                newRow[dayWord] = 1;
                            }


                            newRow["MTD"] = monthEntry.Value.Dates.Count;
                            newRow["FromDate"] = fromDateRow.ToString("dd-MM-yyyy");
                            newRow["ToDate"] = toDateRow.ToString("dd-MM-yyyy");

                            // Add the row to the DataTable
                            dummypnaposTable.Rows.Add(newRow);
                        }

                    }




                    // Call the method to save filtered data to SharePoint list
                    SaveFiltesredDataToSharePointListInBatches(siteUrlHR, username, password, "TravelDates", dummyDataTable);
                    SaveFiltesredDataToSharePointListInBatches(siteUrlHR, username, password, "TravelDates", dummypnaposTable);
                    
                }
            }
        }

        static void ConvertCsvToXlsx(string csvFilePath, string xlsxFilePath)
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Load CSV data
                string[] csvLines = File.ReadAllLines(csvFilePath);

                // Write CSV data to Excel worksheet
                for (int i = 0; i < csvLines.Length; i++)
                {
                    string[] fields = SplitCsvLine(csvLines[i]);

                    for (int j = 0; j < fields.Length; j++)
                    {
                        // Write each field to the corresponding cell in Excel
                        worksheet.Cells[i + 1, j + 1].Value = fields[j];
                    }
                }

                // Save XLSX file
                FileInfo xlsxFile = new FileInfo(xlsxFilePath);
                package.SaveAs(xlsxFile);
            }
        }

        static string[] SplitCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            int startIndex = 0;

            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (line[i] == ',' && !inQuotes)
                {
                    fields.Add(line.Substring(startIndex, i - startIndex).Trim('"'));
                    startIndex = i + 1;
                }
            }

            // Add the last field
            fields.Add(line.Substring(startIndex).Trim('"'));

            return fields.ToArray();
        }


   private static void SaveFiltesredDataToSharePointListInBatches(string siteUrl, string username, string password, string listTitle, DataTable dataTable)
        {
            // Connect to SharePoint site using credentials
            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            using (var clientContext = new Microsoft.SharePoint.Client.ClientContext(siteUrl))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(username, securePassword);

                try
                {
                    // Get the SharePoint list
                    List list = clientContext.Web.Lists.GetByTitle(listTitle);
                    clientContext.Load(list);
                    clientContext.ExecuteQuery();

                    // Include header row if needed
                    DataRow headerRow = dataTable.Rows[0];
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                    // Define batch size for uploading data
                    int batchSize = 1; // Adjust batch size as needed
                    int totalItems = dataTable.Rows.Count;

                    for (int i = 0; i < dataTable.Rows.Count; i += batchSize)
                    {
                        var batch = new List<ListItem>();

                        // Process the current batch
                        for (int j = i; j < Math.Min(i + batchSize, dataTable.Rows.Count); j++)
                        {
                            DataRow row = dataTable.Rows[j];
                            ListItem newItem = null; // Define newItem here

                            // Check if an item with the same Country, Company, and Month already exists in the SharePoint list
                            string country = row["Country"].ToString();
                            string company = row["Company"].ToString();
                            string month = row["Month"].ToString();
                            var query = new CamlQuery();
                            query.ViewXml = $@"<View>
                       <Query>
                         <Where>
                           <And>
                             <And>
                               <Eq>
                                 <FieldRef Name='Country' />
                                 <Value Type='Text'>{country}</Value>
                               </Eq>
                               <Eq>
                                 <FieldRef Name='Company' />
                                 <Value Type='Text'>{company}</Value>
                               </Eq>
                             </And>
                             <Eq>
                               <FieldRef Name='Month' />
                               <Value Type='Text'>{month}</Value>
                             </Eq>
                           </And>
                         </Where>
                       </Query>
                     </View>";
                            var items = list.GetItems(query);
                            clientContext.Load(items);
                            clientContext.ExecuteQuery();

                            // Calculate the MTD count for filtered records
                            //int totalMTDCount = 0;


                            ListItem existingItem = items.FirstOrDefault();

                            if (existingItem != null)
                            {
                                // Update existing item
                                foreach (DataColumn col in dataTable.Columns)
                                {
                                    string columnName = MapToSharePointColumnName(col.ColumnName);
                                    object columnValue = row[col];

                                    // Check if the cell value is 1, then update the corresponding SharePoint list item
                                    if (columnValue.ToString() == "1")
                                    {
                                        existingItem[columnName] = columnValue;
                                    }
                                    if (columnName == "MTD")
                                    {
                                        existingItem[columnName] = columnValue;
                                        Console.WriteLine("MTD column updated successfully");
                                    }
                                    // Increment the item's version to indicate changes
                                       }
                                existingItem.Update();
                            }

                            else
                            {
                                // Create new list item
                                newItem = list.AddItem(itemCreateInfo);
                                foreach (DataColumn col in dataTable.Columns)
                                {
                                    string columnName = MapToSharePointColumnName(col.ColumnName);
                                    object columnValue = row[col];

                                    // Check if the cell value is 1, then update the corresponding SharePoint list item
                                    //if (columnValue.ToString() == "1")
                                    //{
                                    newItem[columnName] = columnValue;
                                    // }
                                    if (columnName == "MTD")
                                    {
                                        newItem[columnName] = columnValue;
                                        Console.WriteLine("MTD column updated successfully");
                                    }
                                }
                                newItem.Update();

                            }


                            batch.Add(existingItem != null ? existingItem : newItem);

                        }

                        // Save the changes for the current batch
                        // Save the changes for the current batch
                        foreach (var item in batch)
                        {
                            try
                            {
                                item.Update();
                                clientContext.ExecuteQuery();
                            }
                            catch (Microsoft.SharePoint.Client.ServerException ex)
                            {
                                // Log the exception
                                Console.WriteLine($"Error occurred while saving data to SharePoint list: {ex.Message}");
                                // Optionally, you can throw the exception to propagate it further
                                Console.WriteLine($"Retrying the operation after a delay...");
                                System.Threading.Thread.Sleep(1000); // Introduce a delay of 1 second (adjust as needed)
                                item.Update(); // Retry updating the item
                                clientContext.ExecuteQuery(); // Retry executing the query
                            }
                        }


                        Console.WriteLine($"Batch processed. {Math.Min(totalItems, i + batchSize)} items added.");
                    }

                    string[] numberWords = { "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty", "Twentyone", "Twentytwo", "Twentythree", "Twentyfour", "Twentyfive", "Twentysix", "Twentyseven", "Twentyeight", "Twentynine", "Thirty", "Thirtyone" };

                    // Update the MTD column based on the count of "1s" in the rows
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string country = row["Country"].ToString();
                        string company = row["Company"].ToString();
                        string month = row["Month"].ToString();

                        var query = new CamlQuery();
                        query.ViewXml = $@"<View>
                                   <Query>
                                     <Where>
                                       <And>
                                         <And>
                                           <Eq>
                                             <FieldRef Name='Country' />
                                             <Value Type='Text'>{country}</Value>
                                           </Eq>
                                           <Eq>
                                             <FieldRef Name='Company' />
                                             <Value Type='Text'>{company}</Value>
                                           </Eq>
                                         </And>
                                         <Eq>
                                           <FieldRef Name='Month' />
                                           <Value Type='Text'>{month}</Value>
                                         </Eq>
                                       </And>
                                     </Where>
                                   </Query>
                                 </View>";

                        var items = list.GetItems(query);
                        clientContext.Load(items);
                        clientContext.ExecuteQuery();

                        foreach (var item in items)
                        {
                            int mtdCount = 0;
                            foreach (string word in numberWords)
                            {
                                if (item[word] != null && item[word].ToString() == "1")
                                {
                                    mtdCount++;
                                }
                            }
                            item["MTD"] = mtdCount;
                            item.Update();
                        }

                        clientContext.ExecuteQuery();
                    }

                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                }
                catch (Exception ex)
                {
                    // Log the exception
                    Console.WriteLine($"Error occurred while saving data to SharePoint list: {ex.Message}");
                    // Optionally, you can throw the exception to propagate it further
                    throw;
                }
            }
        }



        private static string MapToSharePointColumnName(string excelColumnName)
        {
            // Map Excel column names to SharePoint column names
            switch (excelColumnName)
            {
                case "title":
                    return "Title";
                case "date":
                    return "Title";
                // Additional mappings go here
                default:
                    return excelColumnName;
            }
        }



       
    }
}