using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Threading;
using System.Net;
using System.Data;
using static JSE_Reports.Definition.Domain;
using System.Configuration;
using System.Globalization;
using System.Text.RegularExpressions;

namespace JSE_Reports.Repository
{
    public class JSEReports
    {
        public class FileDownloader
        {
            public void DownloadFiles(string targetDirectory)
            {
                var baseUrl = "https://clientportal.jse.co.za";
                var pageUrl = $"{baseUrl}/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";
                var web = new HtmlWeb();
                var doc = web.Load(pageUrl);

                //XPath query
                var links = doc.DocumentNode.SelectNodes("//a[contains(@href, '_D_Daily MTM Report') and contains(@href, '2023') and contains(@href, '.xls')]")
                    .Select(node => node.GetAttributeValue("href", ""))
                    .Distinct();


                using (var webClient = new WebClient())
                {
                    foreach (var link in links)
                    {
                        var fileUrl = baseUrl + link;
                        var fileName = Path.GetFileName(fileUrl);
                        var localFilePath = Path.Combine(targetDirectory, fileName);

                        if (!File.Exists(localFilePath))
                        {
                            Console.WriteLine($"Downloading file: {fileName}");
                            webClient.DownloadFile(fileUrl, localFilePath);
                        }
                        else
                        {
                            Console.WriteLine($"File already exists: {fileName}");
                        }
                    }
                }
            }
        }
        public class FileUploader
        {
            public void ProcessExcelFile(string filePath)
            {
                var mtmDataList = ReadExcelData(filePath);
                var mtmDataTable = ToDataTable(mtmDataList);
                InsertDataIntoDatabase(mtmDataTable);
            }

            public List<DailyMTM> ReadExcelData(string filePath)
            {
                List<DailyMTM> mtmDataList = new List<DailyMTM>();
                Application excelApp = null;
                Workbook workbook = null;
                try
                {
                    excelApp = new Application();
                    workbook = excelApp.Workbooks.Open(filePath);
                    Worksheet worksheet = workbook.Sheets[1];
                    Range range = worksheet.UsedRange;

                    // Extracting FileDate from cell:(row 3, column 1)
                    string dateCellContent = (range.Cells[3, 1] as Range).Text;
                    DateTime fileDate = ExtractDateFromString(dateCellContent);

                    for (int row = 6; row <= range.Rows.Count; row++)
                    {
                        DailyMTM mtmData = new DailyMTM
                        {
                            FileDate = fileDate,
                            Contract = Convert.ToString(range.Cells[row, 1].Value2),
                            ExpiryDate = ConvertToDate(range.Cells[row, 3].Value2),
                            Classification = range.Cells[row, 4].Value2 != null ? Convert.ToString(range.Cells[row, 4].Value2) : string.Empty,
                            Strike = ConvertToDouble(range.Cells[row, 5].Value2),
                            CallPut = Convert.ToString(range.Cells[row, 6].Value2),  
                            MTMYield = Convert.ToDouble(range.Cells[row, 7].Value2),
                            MarkPrice = Convert.ToDouble(range.Cells[row, 8].Value2),
                            SpotRate = Convert.ToDouble(range.Cells[row, 9].Value2),
                            PreviousMTM = Convert.ToDouble(range.Cells[row, 10].Value2),
                            PreviousPrice = Convert.ToDouble(range.Cells[row, 11].Value2),
                            PremiumOnOption = Convert.ToDouble(range.Cells[row, 12].Value2),    
                            Volatility = Convert.ToDouble(range.Cells[row, 13].Value2),
                            Delta = Convert.ToDouble(range.Cells[row, 14].Value2),
                            DeltaValue = Convert.ToDouble(range.Cells[row, 15].Value2),
                            ContractsTraded = Convert.ToDouble(range.Cells[row, 16].Value2),
                            OpenInterest = Convert.ToDouble(range.Cells[row, 17].Value2)
                        };
                        mtmDataList.Add(mtmData);
                    }

                    //workbook.Close(false);
                    //excelApp.Quit();
                    //Marshal.ReleaseComObject(workbook);
                    //Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                finally
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }

                return mtmDataList;
            }

            public System.Data.DataTable ToDataTable(List<DailyMTM> mtmDataList)
            {
                System.Data.DataTable table = new System.Data.DataTable();
                table.Columns.Add("FileDate", typeof(DateTime));
                table.Columns.Add("Contract", typeof(string));
                table.Columns.Add("ExpiryDate", typeof(DateTime));
                table.Columns.Add("Classification", typeof(string));
                table.Columns.Add("Strike", typeof(double));
                table.Columns.Add("CallPut", typeof(string));
                table.Columns.Add("MTMYield", typeof(double));
                table.Columns.Add("MarkPrice", typeof(double));
                table.Columns.Add("SpotRate", typeof(double));
                table.Columns.Add("PreviousMTM", typeof(double));
                table.Columns.Add("PreviousPrice", typeof(double));
                table.Columns.Add("PremiumOnOption", typeof(double));
                table.Columns.Add("Volatility", typeof(double));
                table.Columns.Add("Delta", typeof(double));
                table.Columns.Add("DeltaValue", typeof(double));
                table.Columns.Add("ContractsTraded", typeof(double));
                table.Columns.Add("OpenInterest", typeof(double));

                foreach (var mtm in mtmDataList)
                {
                    DataRow row = table.NewRow();
                    row["FileDate"] = mtm.FileDate;
                    row["Contract"] = mtm.Contract;
                    row["ExpiryDate"] = mtm.ExpiryDate;
                    row["Classification"] = mtm.Classification;
                    row["Strike"] = mtm.Strike;
                    row["CallPut"] = mtm.CallPut;
                    row["MTMYield"] = mtm.MTMYield;
                    row["MarkPrice"] = mtm.MarkPrice;
                    row["SpotRate"] = mtm.SpotRate;
                    row["PreviousMTM"] = mtm.PreviousMTM;
                    row["PreviousPrice"] = mtm.PreviousPrice;
                    row["PremiumOnOption"] = mtm.PremiumOnOption;
                    row["Volatility"] = mtm.Volatility;
                    row["Delta"] = mtm.Delta;
                    row["DeltaValue"] = mtm.DeltaValue;
                    row["ContractsTraded"] = mtm.ContractsTraded;
                    row["OpenInterest"] = mtm.OpenInterest;

                    table.Rows.Add(row);
                }

                return table;
            }

            public void InsertDataIntoDatabase(System.Data.DataTable mtmDataTable)
            {
                try
                {
                    using (var conn = new SqlConnection())
                    using (var cmd = new SqlCommand())
                    using (var da = new SqlDataAdapter(cmd))
                    using (var dt = new System.Data.DataTable())
                    {
                        // Check WebConfig file to update the connection string
                        conn.ConnectionString = ConfigurationManager.ConnectionStrings["ApplicationServices"].ConnectionString;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "InsertDailyMTMData";
                        cmd.Parameters.AddWithValue("@DailyMTMData", mtmDataTable);
                        cmd.Connection = conn;
                        conn.Open();
                        da.Fill(dt);
                        conn.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }

            //// These are helper functions
            public static DateTime ExtractDateFromString(string dateString)
            {
                try
                {
                    string pattern = "DAILY SUMMARY FOR: ";
                    int startIndex = dateString.IndexOf(pattern);
                    if (startIndex == -1)
                    {
                        throw new FormatException("Date pattern not found in the string.");
                    }

                    startIndex += pattern.Length;
                    string datePart = dateString.Substring(startIndex).Trim();
                    string datestring = datePart;

                    DateTime parsedDate = DateTime.ParseExact(datestring, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string formattedDate = parsedDate.ToString("yyyy/MM/dd");

                    return DateTime.ParseExact(formattedDate, "yyyy/MM/dd", CultureInfo.InvariantCulture);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error parsing date: " + ex.Message);
                    throw;
                }
            }
            private static DateTime ConvertToDate(object value)
            {
                if (value == null || value is DBNull)
                {
                    return default(DateTime); 
                }
                return DateTime.FromOADate(Convert.ToDouble(value));
            }
            private static double ConvertToDouble(object value)
            {
                if (value == null || value is DBNull)
                {
                    return 0.0;
                }
                return Convert.ToDouble(value);
            }
        }
        public class FileUploaderAsync
        {
            public async Task ProcessExcelFile(string filePath)
            {
                var mtmDataList = await ReadExcelData(filePath);
                var mtmDataTable = ToDataTable(mtmDataList);
                await InsertDataIntoDatabase(mtmDataTable);
            }

            public async Task<List<DailyMTM>> ReadExcelData(string filePath)
            {

                return await Task.Run(() =>
                {
                    List<DailyMTM> mtmDataList = new List<DailyMTM>();
                    Application excelApp = null;
                    Workbook workbook = null;
                    try
                    {
                        excelApp = new Application();
                        workbook = excelApp.Workbooks.Open(filePath);
                        Worksheet worksheet = workbook.Sheets[1];
                        Range range = worksheet.UsedRange;

                        // Extracting FileDate from cell:(row 3, column 1)
                        string dateCellContent = (range.Cells[3, 1] as Range).Text;
                        DateTime fileDate = ExtractDateFromString(dateCellContent);

                        for (int row = 6; row <= range.Rows.Count; row++)
                        {
                            DailyMTM mtmData = new DailyMTM
                            {
                                FileDate = fileDate,
                                Contract = Convert.ToString(range.Cells[row, 1].Value2),
                                ExpiryDate = ConvertToDate(range.Cells[row, 3].Value2),
                                Classification = range.Cells[row, 4].Value2 != null ? Convert.ToString(range.Cells[row, 4].Value2) : string.Empty,
                                Strike = ConvertToDouble(range.Cells[row, 5].Value2),
                                CallPut = Convert.ToString(range.Cells[row, 6].Value2),  
                                MTMYield = Convert.ToDouble(range.Cells[row, 7].Value2),
                                MarkPrice = Convert.ToDouble(range.Cells[row, 8].Value2),
                                SpotRate = Convert.ToDouble(range.Cells[row, 9].Value2),
                                PreviousMTM = Convert.ToDouble(range.Cells[row, 10].Value2),
                                PreviousPrice = Convert.ToDouble(range.Cells[row, 11].Value2),
                                PremiumOnOption = Convert.ToDouble(range.Cells[row, 12].Value2),    
                                Volatility = Convert.ToDouble(range.Cells[row, 13].Value2),
                                Delta = Convert.ToDouble(range.Cells[row, 14].Value2),
                                DeltaValue = Convert.ToDouble(range.Cells[row, 15].Value2),
                                ContractsTraded = Convert.ToDouble(range.Cells[row, 16].Value2),
                                OpenInterest = Convert.ToDouble(range.Cells[row, 17].Value2)
                            };
                            mtmDataList.Add(mtmData);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                    }

                    return mtmDataList;
                });
            }

            public System.Data.DataTable ToDataTable(List<DailyMTM> mtmDataList)
            {
                System.Data.DataTable table = new System.Data.DataTable();
                table.Columns.Add("FileDate", typeof(DateTime));
                table.Columns.Add("Contract", typeof(string));
                table.Columns.Add("ExpiryDate", typeof(DateTime));
                table.Columns.Add("Classification", typeof(string));
                table.Columns.Add("Strike", typeof(double));
                table.Columns.Add("CallPut", typeof(string));
                table.Columns.Add("MTMYield", typeof(double));
                table.Columns.Add("MarkPrice", typeof(double));
                table.Columns.Add("SpotRate", typeof(double));
                table.Columns.Add("PreviousMTM", typeof(double));
                table.Columns.Add("PreviousPrice", typeof(double));
                table.Columns.Add("PremiumOnOption", typeof(double));
                table.Columns.Add("Volatility", typeof(double));
                table.Columns.Add("Delta", typeof(double));
                table.Columns.Add("DeltaValue", typeof(double));
                table.Columns.Add("ContractsTraded", typeof(double));
                table.Columns.Add("OpenInterest", typeof(double));

                foreach (var mtm in mtmDataList)
                {
                    DataRow row = table.NewRow();
                    row["FileDate"] = mtm.FileDate;
                    row["Contract"] = mtm.Contract;
                    row["ExpiryDate"] = mtm.ExpiryDate;
                    row["Classification"] = mtm.Classification;
                    row["Strike"] = mtm.Strike;
                    row["CallPut"] = mtm.CallPut;
                    row["MTMYield"] = mtm.MTMYield;
                    row["MarkPrice"] = mtm.MarkPrice;
                    row["SpotRate"] = mtm.SpotRate;
                    row["PreviousMTM"] = mtm.PreviousMTM;
                    row["PreviousPrice"] = mtm.PreviousPrice;
                    row["PremiumOnOption"] = mtm.PremiumOnOption;
                    row["Volatility"] = mtm.Volatility;
                    row["Delta"] = mtm.Delta;
                    row["DeltaValue"] = mtm.DeltaValue;
                    row["ContractsTraded"] = mtm.ContractsTraded;
                    row["OpenInterest"] = mtm.OpenInterest;

                    table.Rows.Add(row);
                }

                return table;
            }

            public async Task InsertDataIntoDatabase(System.Data.DataTable mtmDataTable)
            {
                try
                {
                    using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ApplicationServices"].ConnectionString))
                    using (var cmd = new SqlCommand("InsertDailyMTMData", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@DailyMTMData", mtmDataTable);

                        await conn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }

            //// These are helper functions
            public static DateTime ExtractDateFromString(string dateString)
            {
                try
                {
                    string pattern = "DAILY SUMMARY FOR: ";
                    int startIndex = dateString.IndexOf(pattern);
                    if (startIndex == -1)
                    {
                        throw new FormatException("Date pattern not found in the string.");
                    }

                    startIndex += pattern.Length;
                    string datePart = dateString.Substring(startIndex).Trim();
                    string datestring = datePart;

                   
                    DateTime parsedDate = DateTime.ParseExact(datestring, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string formattedDate = parsedDate.ToString("yyyy/MM/dd");

                    return DateTime.ParseExact(formattedDate, "yyyy/MM/dd", CultureInfo.InvariantCulture);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error parsing date: " + ex.Message);
                    throw;
                }
            }
            private static DateTime ConvertToDate(object value)
            {
                if (value == null || value is DBNull)
                {
                    return default(DateTime);
                }
                return DateTime.FromOADate(Convert.ToDouble(value));
            }
            private static double ConvertToDouble(object value)
            {
                if (value == null || value is DBNull)
                {
                    return 0.0;
                }
                return Convert.ToDouble(value);
            }
        }
    }
}