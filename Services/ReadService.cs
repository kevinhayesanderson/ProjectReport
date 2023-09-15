using ExcelDataReader;
using Models;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.Json;
using Utilities;

namespace Services
{
    public static class ReadService
    {
        private static readonly List<object> BookingMonths = new();

        private static readonly double[] Months = new double[12] { 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0 };
        
        private static int _ptrBookingMonthCol;

        public static MonthlyReportData ReadMonthlyReports(List<string> monthlyReports, UserSettings userSettings)
        {
            List<EmployeeData> employeeDataList = new();
            foreach (string monthlyReport in monthlyReports)
            {
                try
                {
                    ConsoleLogger.Log("Reading " + new FileInfo(monthlyReport).Name, 2);
                    using (FileStream fileStream = File.Open(monthlyReport, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null))
                        {
                            DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                            List<DataTable> dataTableList = userSettings.MonthlyReportMonths.Any() ? tables.Cast<DataTable>().Where(dataTable => userSettings.MonthlyReportMonths.Contains(dataTable.TableName.Trim())).ToList() : tables.Cast<DataTable>().Select(dataTable => dataTable).ToList();
                            if (dataTableList.Count > 0)
                            {
                                if (string.IsNullOrEmpty(((string)dataTableList[0].Rows[3][2]).Trim()))
                                {
                                    DefaultInterpolatedStringHandler interpolatedStringHandler = new(68, 3);
                                    interpolatedStringHandler.AppendLiteral("Name cannot be null or empty string in sheet ");
                                    interpolatedStringHandler.AppendFormatted(dataTableList[0].TableName);
                                    interpolatedStringHandler.AppendLiteral(": Check row ");
                                    interpolatedStringHandler.AppendFormatted(4);
                                    interpolatedStringHandler.AppendLiteral(" at column ");
                                    interpolatedStringHandler.AppendFormatted(3);
                                    throw new ArgumentException(interpolatedStringHandler.ToStringAndClear());
                                }
                                string str = (string)dataTableList[0].Rows[3][2];
                                int result1;
                                if (!int.TryParse(dataTableList[0].Rows[4][2].ToString(), out result1))
                                {
                                    DefaultInterpolatedStringHandler interpolatedStringHandler = new(47, 3);
                                    interpolatedStringHandler.AppendLiteral("Invalid format at column ");
                                    interpolatedStringHandler.AppendFormatted(3);
                                    interpolatedStringHandler.AppendLiteral(": Check row ");
                                    interpolatedStringHandler.AppendFormatted(5);
                                    interpolatedStringHandler.AppendLiteral(" in sheet ");
                                    interpolatedStringHandler.AppendFormatted(dataTableList[0].TableName);
                                    throw new FormatException(interpolatedStringHandler.ToStringAndClear());
                                }
                                Dictionary<string, TimeSpan> dictionary = new();
                                Console.Write("Reading Sheet ");
                                foreach (DataTable dataTable in dataTableList)
                                {
                                    try
                                    {
                                        Console.Write(dataTable.TableName + ",");
                                        DataRowCollection rows = dataTable.Rows;
                                        for (int index = 15; index < rows.Count; ++index)
                                        {
                                            if (rows[index][2] is string key && (key.ToUpper().StartsWith("ACS_", StringComparison.Ordinal) || key.ToUpper().StartsWith("ACS.", StringComparison.Ordinal)))
                                            {
                                                if (key.ToUpper().StartsWith("ACS.", StringComparison.Ordinal))
                                                    key = key.Replace('.', '_');
                                                int columnIndex = rows[index].ItemArray.Length - 1;
                                                TimeSpan timeSpan;
                                                if (dictionary.TryGetValue(key, out timeSpan))
                                                {
                                                    TimeSpan result2;
                                                    if (TimeSpan.TryParse(rows[index][columnIndex].ToString(), out result2))
                                                    {
                                                        if (result2.Equals(new TimeSpan()))
                                                        {
                                                            DefaultInterpolatedStringHandler interpolatedStringHandler = new(67, 3);
                                                            interpolatedStringHandler.AppendLiteral("Check for potential data mismatch at column ");
                                                            interpolatedStringHandler.AppendFormatted(columnIndex + 1);
                                                            interpolatedStringHandler.AppendLiteral(": Check row ");
                                                            interpolatedStringHandler.AppendFormatted(index + 1);
                                                            interpolatedStringHandler.AppendLiteral(" in sheet ");
                                                            interpolatedStringHandler.AppendFormatted(dataTable.TableName);
                                                            interpolatedStringHandler.AppendLiteral(".");
                                                            ConsoleLogger.LogWarning(interpolatedStringHandler.ToStringAndClear(), 1);
                                                        }
                                                        timeSpan += result2;
                                                        dictionary[key] = timeSpan;
                                                    }
                                                    else
                                                    {
                                                        DefaultInterpolatedStringHandler interpolatedStringHandler = new(48, 3);
                                                        interpolatedStringHandler.AppendLiteral("Invalid format at column ");
                                                        interpolatedStringHandler.AppendFormatted(columnIndex + 1);
                                                        interpolatedStringHandler.AppendLiteral(": Check row ");
                                                        interpolatedStringHandler.AppendFormatted(index + 1);
                                                        interpolatedStringHandler.AppendLiteral(" in sheet ");
                                                        interpolatedStringHandler.AppendFormatted(dataTable.TableName);
                                                        interpolatedStringHandler.AppendLiteral(".");
                                                        throw new FormatException(interpolatedStringHandler.ToStringAndClear());
                                                    }
                                                }
                                                else
                                                {
                                                    TimeSpan result3;
                                                    if (TimeSpan.TryParse(rows[index][columnIndex].ToString(), out result3))
                                                    {
                                                        if (result3.Equals(new TimeSpan()))
                                                        {
                                                            DefaultInterpolatedStringHandler interpolatedStringHandler = new(67, 3);
                                                            interpolatedStringHandler.AppendLiteral("Check for potential data mismatch at column ");
                                                            interpolatedStringHandler.AppendFormatted(columnIndex + 1);
                                                            interpolatedStringHandler.AppendLiteral(": Check row ");
                                                            interpolatedStringHandler.AppendFormatted(index + 1);
                                                            interpolatedStringHandler.AppendLiteral(" in sheet ");
                                                            interpolatedStringHandler.AppendFormatted(dataTable.TableName);
                                                            interpolatedStringHandler.AppendLiteral(".");
                                                            ConsoleLogger.LogWarning(interpolatedStringHandler.ToStringAndClear(), 1);
                                                        }
                                                        dictionary[key] = result3;
                                                    }
                                                    else
                                                    {
                                                        DefaultInterpolatedStringHandler interpolatedStringHandler = new(48, 3);
                                                        interpolatedStringHandler.AppendLiteral("Invalid format at column ");
                                                        interpolatedStringHandler.AppendFormatted(columnIndex + 1);
                                                        interpolatedStringHandler.AppendLiteral(": Check row ");
                                                        interpolatedStringHandler.AppendFormatted(index + 1);
                                                        interpolatedStringHandler.AppendLiteral(" in sheet ");
                                                        interpolatedStringHandler.AppendFormatted(dataTable.TableName);
                                                        interpolatedStringHandler.AppendLiteral(".");
                                                        throw new FormatException(interpolatedStringHandler.ToStringAndClear());
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ConsoleLogger.LogErrorAndExit("Error on reading sheet " + dataTable.TableName + ": " + ex.Message);
                                    }
                                }
                                employeeDataList.Add(new EmployeeData()
                                {
                                    Id = result1,
                                    Name = str,
                                    ProjectTime = dictionary
                                });
                            }
                            else
                                ConsoleLogger.LogWarning("No sheets found for the filter condition in monthly report " + monthlyReport + ".");
                        }
                    }
                }
                catch (Exception ex)
                {
                    ConsoleLogger.LogErrorAndExit("Error on reading monthly report " + monthlyReport + ": " + ex.Message);
                }
            }
            return new MonthlyReportData()
            {
                EmployeesData = employeeDataList
            };
        }

        public static PtrData ReadPtr(List<string> ptrFiles, UserSettings userSettings)
        {
            PtrData ptrData = new();
            _ptrBookingMonthCol = userSettings.PtrBookingMonthCol;
            bool flag = userSettings.PtrBookingMonths.Count == 0;
            if (!flag)
                ConvertInputBookingMonths(userSettings.PtrBookingMonths);
            ConsoleLogger.LogLine();
            try
            {
                Dictionary<string, double> projectEfforts = new();
                foreach (string ptrFile in ptrFiles)
                {
                    ConsoleLogger.Log("Reading " + new FileInfo(ptrFile).Name);
                    using (FileStream fileStream = File.Open(ptrFile, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null))
                        {
                            DataTable table = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables[userSettings.PtrSheetName];
                            if (table == null)
                            {
                                ConsoleLogger.LogErrorAndExit("Error on reading Ptr: no sheet found for name " + userSettings.PtrSheetName);
                            }
                            else
                            {
                                List<DataRow> list = table.AsEnumerable().Where(r => r[userSettings.PtrBookingMonthCol - 1] != DBNull.Value).Skip(1).ToList();
                                List<DataRow> rows = flag ? list : list.Where(new Func<DataRow, bool>(IsRowOfBookingMonth)).ToList();
                                int projectIdCol = userSettings.PtrProjectIdCol - 1;
                                if (rows.Any())
                                {
                                    bool isEffortInDouble = false;
                                    bool isEffortInTimeSpan = false;
                                    userSettings.PtrEffortCols.ForEach(ef =>
                                    {
                                        object obj = rows.First()[(int)ef - 1];
                                        isEffortInDouble = obj is double;
                                        isEffortInTimeSpan = obj is TimeSpan;
                                        if (isEffortInDouble || isEffortInTimeSpan)
                                            return;

                                        DefaultInterpolatedStringHandler interpolatedStringHandler = new(50, 1);
                                        interpolatedStringHandler.AppendLiteral("Unknown format of the effort column values: column");
                                        interpolatedStringHandler.AppendFormatted(ef);
                                        ConsoleLogger.LogErrorAndExit(interpolatedStringHandler.ToStringAndClear());
                                    });
                                    rows.ForEach(row =>
                                    {
                                        if (!((string)row[projectIdCol]).ToUpper().StartsWith("ACS", StringComparison.Ordinal))
                                            return;
                                        string key = ((string)row[projectIdCol]).Replace('.', '_');
                                        double totalEffort = 0.0;
                                        userSettings.PtrEffortCols.ForEach(ef =>
                                        {
                                            if (row[(int)ef - 1] == DBNull.Value)
                                                totalEffort += 0.0;
                                            if (isEffortInDouble)
                                                totalEffort += Convert.ToDouble(row[(int)ef - 1], CultureInfo.InvariantCulture);
                                            if (!isEffortInTimeSpan)
                                                return;
                                            totalEffort += ((TimeSpan)row[(int)ef - 1]).TotalHours;
                                        });
                                        if (projectEfforts.ContainsKey(key))
                                            projectEfforts[key] += totalEffort;
                                        else
                                            projectEfforts[key] = totalEffort;
                                    });
                                }
                            }
                        }
                    }
                }
                ptrData = new PtrData()
                {
                    ProjectEfforts = projectEfforts
                };
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on reading Ptr: " + ex.Message);
            }
            return ptrData;
        }

        public static UserSettings ReadUserSettings()
        {
            UserSettings userSettings1 = new();
            try
            {
                ConsoleLogger.LogInfo("Reading user settings.");
                UserSettings userSettings2 = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText("userSettings.json"));
                if (userSettings2 != null)
                {
                    userSettings1 = userSettings2;
                    ConsoleLogger.LogInfo("Read settings:");
                    ConsoleLogger.Log("Folder:" + userSettings1.Folder);
                    ConsoleLogger.Log("MonthlyReportMonths:" + string.Join(",", (IEnumerable<string>)userSettings1.MonthlyReportMonths));
                    ConsoleLogger.Log("PtrSheetName:" + userSettings1.PtrSheetName);
                    DefaultInterpolatedStringHandler interpolatedStringHandler1 = new(1, 2);
                    interpolatedStringHandler1.AppendFormatted("PtrProjectIdCol");
                    interpolatedStringHandler1.AppendLiteral(":");
                    interpolatedStringHandler1.AppendFormatted(userSettings1.PtrProjectIdCol);
                    ConsoleLogger.Log(interpolatedStringHandler1.ToStringAndClear());
                    DefaultInterpolatedStringHandler interpolatedStringHandler2 = new(1, 2);
                    interpolatedStringHandler2.AppendFormatted("PtrBookingMonthCol");
                    interpolatedStringHandler2.AppendLiteral(":");
                    interpolatedStringHandler2.AppendFormatted(userSettings1.PtrBookingMonthCol);
                    ConsoleLogger.Log(interpolatedStringHandler2.ToStringAndClear());
                    ConsoleLogger.Log("PtrBookingMonths:" + string.Join(",", userSettings1.PtrBookingMonths.Select(bookingMonth => bookingMonth).ToArray()));
                    ConsoleLogger.Log("PtrEffortCols:" + string.Join(",", userSettings1.PtrEffortCols.Select(effortCol => effortCol.ToString(CultureInfo.InvariantCulture)).ToArray()));
                    interpolatedStringHandler2 = new DefaultInterpolatedStringHandler(1, 2);
                    interpolatedStringHandler2.AppendFormatted("GenerateLeaveReport");
                    interpolatedStringHandler2.AppendLiteral(":");
                    interpolatedStringHandler2.AppendFormatted(userSettings1.GenerateLeaveReport);
                    ConsoleLogger.Log(interpolatedStringHandler2.ToStringAndClear());
                    ConsoleLogger.Log("FinancialYear:" + userSettings1.FinancialYear);
                }
                else
                    ConsoleLogger.LogErrorAndExit("Error on reading user settings: null userSettings value");
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on reading user settings: " + ex.Message);
            }
            return userSettings1;
        }

        private static bool ConvertInputBookingMonths(List<object> inputBookingMonths)
        {
            inputBookingMonths.ForEach(ibm =>
            {
                JsonElement jsonElement = (JsonElement)ibm;
                switch (jsonElement.ValueKind)
                {
                    case JsonValueKind.String:
                        string str = jsonElement.GetString();
                        if (string.IsNullOrEmpty(str))
                            break;
                        if (str.Contains('|'))
                        {
                            if (Months.Contains(Convert.ToDouble(str.Split('|')[0].Trim())))
                                BookingMonths.Add(Convert.ToDouble(str.Split('|')[0].Trim()));
                            BookingMonths.Add(str.Split('|')[1].Trim());
                            break;
                        }
                        BookingMonths.Add(str.Trim());
                        break;

                    case JsonValueKind.Number:
                        if (!Months.Contains(jsonElement.GetDouble()))
                            break;
                        BookingMonths.Add(jsonElement.GetDouble());
                        break;
                }
            });
            return true;
        }

        private static bool IsRowOfBookingMonth(DataRow dataRow)
        {
            bool flag1 = false;
            bool flag2;
            switch (dataRow[_ptrBookingMonthCol - 1])
            {
                case double num:
                    flag2 = BookingMonths.Contains(num);
                    break;

                case DateTime dateTime:
                    flag2 = BookingMonths.Contains(dateTime.ToString("MMM-yy"));
                    break;

                default:
                    flag2 = flag1;
                    break;
            }
            return flag2;
        }
    }
}