using ExcelDataReader;
using Models;
using System.Data;
using System.Globalization;
using System.Text.Json;
using Utilities;

namespace Services
{
    public static class ReadService
    {
        private static readonly List<object> BookingMonths = new();

        private static int _ptrBookingMonthCol;

        public static MonthlyReportData ReadMonthlyReports(List<string> monthlyReports, UserSettings userSettings)
        {
            List<EmployeeData> employeeDataList = new();
            HashSet<string> projectIds = new();
            foreach (string monthlyReport in monthlyReports)
            {
                try
                {
                    ConsoleLogger.LogInfo("Reading " + new FileInfo(monthlyReport).Name, 1);
                    using FileStream fileStream = File.Open(monthlyReport, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> dataTableList;
                    if (userSettings.MonthlyReportMonths.Any())
                    {
                        dataTableList = tables.Cast<DataTable>().Where(dataTable => userSettings.MonthlyReportMonths.Contains(dataTable.TableName.Trim())).ToList();
                    }
                    else
                    {
                        dataTableList = tables.Cast<DataTable>().Select(dataTable => dataTable).ToList();
                    }
                    if (dataTableList.Count > 0)
                    {
                        if (string.IsNullOrEmpty(((string)dataTableList[0].Rows[3][2]).Trim()))
                        {
                            throw new ArgumentException("Employee name is empty or has an invalid format in the sheet " + dataTableList[0].TableName + ": Check row " + 4 + " at column " + 3);
                        }
                        string employeeName = (string)dataTableList[0].Rows[3][2];
                        employeeName = employeeName.Trim();
                        if (!int.TryParse(dataTableList[0].Rows[4][2].ToString(), out int employeeId))
                        {
                            throw new ArgumentException("Employee Id is empty or has an invalid format in the sheet " + dataTableList[0].TableName + ": Check row " + 5 + " at column " + 3);
                        }
                        Dictionary<string, TimeSpan> projectData = new();
                        ConsoleLogger.LogSameLine("Reading Sheet: ");
                        TimeSpan ActualAvailableHours = new();
                        int TotalLeaves = new();
                        foreach (DataTable dataTable in dataTableList)
                        {
                            try
                            {
                                ConsoleLogger.LogDataSameLine(dataTable.TableName + ", ");
                                DataRowCollection rows = dataTable.Rows;
                                int lastColumnIndex;
                                int actualAvailableHoursRowIndex = 13;
                                lastColumnIndex = rows[actualAvailableHoursRowIndex].ItemArray.Length - 1;
                                if (TimeSpan.TryParse(rows[actualAvailableHoursRowIndex][lastColumnIndex].ToString(), out TimeSpan actualAvailableHours))
                                {
                                    if (actualAvailableHours.Equals(new TimeSpan()))
                                    {
                                        ConsoleLogger.LogWarning($"Check for potential empty data at column: {lastColumnIndex + 1} at row: {actualAvailableHoursRowIndex + 1} in sheet {dataTable.TableName}", 1);
                                    }
                                    ActualAvailableHours += actualAvailableHours;
                                }
                                else
                                {
                                    throw new FormatException($"Invalid format at column: {lastColumnIndex + 1} at row: {actualAvailableHoursRowIndex + 1} in sheet {dataTable.TableName}");
                                }
                                int totalLeavesRowIndex = 14;
                                if (int.TryParse(rows[totalLeavesRowIndex][lastColumnIndex].ToString(), out int totalLeaves))
                                {
                                    if (actualAvailableHours.Equals(new int()))
                                    {
                                        ConsoleLogger.LogWarning($"Check for potential empty data at column: {lastColumnIndex + 1} at row: {totalLeavesRowIndex + 1} in sheet {dataTable.TableName}", 1);
                                    }
                                    TotalLeaves += totalLeaves;
                                }
                                else
                                {
                                    throw new FormatException($"Invalid format at column: {lastColumnIndex + 1} at row: {totalLeavesRowIndex + 1} in sheet {dataTable.TableName}");
                                }
                                for (int rowIndex = 15; rowIndex < rows.Count; ++rowIndex)
                                {
                                    if (rows[rowIndex][2] is string key && (key.ToUpper().StartsWith("ACS_", StringComparison.Ordinal) || key.ToUpper().StartsWith("ACS.", StringComparison.Ordinal)))
                                    {
                                        if (key.ToUpper().StartsWith("ACS.", StringComparison.Ordinal))
                                            key = key.Replace('.', '_');
                                        lastColumnIndex = rows[rowIndex].ItemArray.Length - 1;
                                        if (TimeSpan.TryParse(rows[rowIndex][lastColumnIndex].ToString(), out TimeSpan hours))
                                        {
                                            if (hours.Equals(new TimeSpan()))
                                            {
                                                ConsoleLogger.LogWarning($"Check for potential empty data at column: {lastColumnIndex + 1} at row: {rowIndex + 1} in sheet {dataTable.TableName}", 1);
                                            }
                                            if (projectData.TryGetValue(key, out TimeSpan projectHours))
                                            {
                                                projectHours += hours;
                                                projectData[key] = projectHours;
                                            }
                                            else
                                            {
                                                projectData[key] = hours;
                                            }
                                            projectIds.Add(key);
                                        }
                                        else
                                        {
                                            throw new FormatException($"Invalid format at column: {lastColumnIndex + 1} at row: {rowIndex + 1} in sheet {dataTable.TableName}");
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
                            Id = employeeId,
                            Name = employeeName,
                            ProjectData = projectData,
                            ActualAvailableHours = ActualAvailableHours,
                            TotalLeaves = TotalLeaves,
                            TotalProjectHours = projectData.Values.AsEnumerable().Aggregate((a, b) => a + b),
                        });
                    }
                    else
                        ConsoleLogger.LogWarning("No sheets found for the filter condition in monthly report " + monthlyReport + ".");
                }
                catch (Exception ex)
                {
                    ConsoleLogger.LogErrorAndExit("Error on reading monthly report " + monthlyReport + ": " + ex.Message);
                }
            }
            return new MonthlyReportData()
            {
                EmployeesData = employeeDataList,
                ProjectIds = projectIds
            };
        }

        public static PtrData ReadPtr(List<string> ptrFiles, UserSettings userSettings)
        {
            PtrData ptrData = new();
            _ptrBookingMonthCol = userSettings.PtrBookingMonthCol;
            bool flag = userSettings.PtrBookingMonths.Count == 0;
            if (!flag && ConvertInputBookingMonths(userSettings.PtrBookingMonths))
            {
                ConsoleLogger.LogLine();
            }
            else
            {
                ConsoleLogger.LogErrorAndExit($"Error converting Ptr booking months: {userSettings.PtrBookingMonths}");
            }
            try
            {
                Dictionary<string, TimeSpan> projectEfforts = new();
                HashSet<string> projectIds = new();
                foreach (string ptrFile in ptrFiles)
                {
                    ConsoleLogger.LogInfo("Reading " + new FileInfo(ptrFile).Name);
                    using FileStream fileStream = File.Open(ptrFile, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTable? table = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables[userSettings.PtrSheetName];
                    if (table == null)
                    {
                        ConsoleLogger.LogErrorAndExit("Error on reading Ptr: no sheet found for name " + userSettings.PtrSheetName);
                    }
                    else
                    {
                        List<DataRow> list = table.AsEnumerable().Where(r => r[userSettings.PtrBookingMonthCol - 1] != DBNull.Value).Skip(1).ToList();
                        List<DataRow> rows = flag ? list : list.Where(IsRowOfBookingMonth).ToList();
                        int projectIdCol = userSettings.PtrProjectIdCol - 1;
                        if (rows.Any())
                        {
                            rows.Select((r, i) => new { item = r, index = i }).ToList()
                                .ForEach(obj =>
                                {
                                    DataRow row = obj.item;
                                    int rowIndex = obj.index;
                                    if (!((string)row[projectIdCol]).ToUpper().StartsWith("ACS", StringComparison.Ordinal))
                                        return;
                                    string key = ((string)row[projectIdCol]).Replace('.', '_');
                                    TimeSpan totalEffort = new();
                                    userSettings.PtrEffortCols.ForEach(ef =>
                                    {
                                        var effortCol = (int)ef - 1;
                                        if (row[effortCol] == DBNull.Value)
                                            totalEffort += TimeSpan.Zero;
                                        else if (row[effortCol] is double doubleValue)
                                            totalEffort += new TimeSpan((int)doubleValue, 0, 0);
                                        else if (row[effortCol] is TimeSpan timeSpan)
                                            totalEffort += timeSpan;
                                        else if (row[effortCol] is DateTime dateTime)
                                            totalEffort += new TimeSpan(dateTime.Hour, dateTime.Minute, dateTime.Second);
                                        else
                                            ConsoleLogger.LogErrorAndExit($"Unknown format of the effort value at column: {ef} at row: {rowIndex}");
                                    });
                                    if (projectEfforts.ContainsKey(key))
                                        projectEfforts[key] += totalEffort;
                                    else
                                        projectEfforts[key] = totalEffort;
                                    projectIds.Add(key);
                                });
                        }
                    }
                }
                ptrData = new PtrData()
                {
                    ProjectEfforts = projectEfforts,
                    ProjectIds = projectIds,
                };
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on reading Ptr: " + ex.Message);
            }
            return ptrData;
        }

        public static UserSettings? ReadUserSettings()
        {
            UserSettings? userSettings = new();
            try
            {
                userSettings = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText("userSettings.json"));
                if (userSettings != null)
                {
                    ConsoleLogger.LogInfo("User settings:", 1);
                    ConsoleLogger.LogSameLine("Folder: "); ConsoleLogger.LogDataSameLine(userSettings.Folder, 1);
                    ConsoleLogger.LogSameLine("MonthlyReportMonths: "); ConsoleLogger.LogDataSameLine(string.Join(",", (IEnumerable<string>)userSettings.MonthlyReportMonths), 1);
                    ConsoleLogger.LogSameLine("PtrSheetName: "); ConsoleLogger.LogDataSameLine(userSettings.PtrSheetName, 1);
                    ConsoleLogger.LogSameLine("PtrProjectIdCol: "); ConsoleLogger.LogDataSameLine(userSettings.PtrProjectIdCol.ToString(), 1);
                    ConsoleLogger.LogSameLine("PtrBookingMonthCol: "); ConsoleLogger.LogDataSameLine(userSettings.PtrBookingMonthCol.ToString(), 1);
                    ConsoleLogger.LogSameLine("PtrBookingMonths: "); ConsoleLogger.LogDataSameLine(string.Join(",", userSettings.PtrBookingMonths.Select(bookingMonth => bookingMonth).ToArray()), 1);
                    ConsoleLogger.LogSameLine("PtrEffortCols: "); ConsoleLogger.LogDataSameLine(string.Join(",", userSettings.PtrEffortCols.Select(effortCol => effortCol.ToString(CultureInfo.InvariantCulture)).ToArray()), 1);
                    ConsoleLogger.LogSameLine("GenerateLeaveReport: "); ConsoleLogger.LogDataSameLine(userSettings.GenerateLeaveReport.ToString(), 1);
                    ConsoleLogger.LogSameLine("FinancialYear: "); ConsoleLogger.LogDataSameLine(userSettings.FinancialYear, 1);
                }
                else
                    ConsoleLogger.LogErrorAndExit("Error on reading user settings: null userSettings value");
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on reading user settings: " + ex.Message);
            }
            return userSettings;
        }

        private static bool ConvertInputBookingMonths(List<object> inputBookingMonths)
        {
            inputBookingMonths.ForEach(ibm =>
            {
                JsonElement jsonElement = (JsonElement)ibm;
                switch (jsonElement.ValueKind)
                {
                    case JsonValueKind.String:
                        string? str = jsonElement.GetString();
                        if (string.IsNullOrEmpty(str))
                            break;
                        if (str.Contains('|'))
                        {
                            string[] splits = str.Split('|');
                            string numMonth = splits[0].Trim();
                            if (int.TryParse(numMonth, out int month))
                                BookingMonths.Add(month);
                            BookingMonths.AddRange(splits[1..]);
                            break;
                        }
                        BookingMonths.Add(str.Trim());
                        break;

                    case JsonValueKind.Number:
                        if (!DataService.Months.Contains(jsonElement.GetInt32()))
                            break;
                        BookingMonths.Add(jsonElement.GetInt32());
                        break;

                    default:
                        break;
                }
            });
            return true;
        }

        private static bool IsRowOfBookingMonth(DataRow dataRow) => dataRow[_ptrBookingMonthCol - 1] switch
        {
            double num => BookingMonths.Contains((int)num),
            string str => BookingMonths.Contains(str),
            _ => false,
        };
    }
}