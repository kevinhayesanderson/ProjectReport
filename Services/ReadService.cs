using ExcelDataReader;
using Models;
using System.Data;
using System.Text.Json;
using Utilities;

namespace Services
{
    public static class ReadService
    {
        private static readonly List<object> BookingMonths = new();

        private static int _ptrBookingMonthCol;

        public static MonthlyReportData ReadMonthlyReports(List<string> reports, object[] months, int idCol)
        {
            List<EmployeeData> employeeDataList = new();
            HashSet<string> projectIds = new();
            foreach (string report in reports)
            {
                try
                {
                    ConsoleLogger.LogInfo("Reading " + new FileInfo(report).Name, 1);
                    using FileStream fileStream = File.Open(report, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> dataTableList = months.Any()
                        ? tables.Cast<DataTable>().Where(dataTable => months.Contains(dataTable.TableName.Trim())).ToList()
                        : tables.Cast<DataTable>().Select(dataTable => dataTable).ToList();
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
                        TimeSpan ActualAvailableHours = TimeSpan.Zero;
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
                                    if (rows[rowIndex][idCol - 1] is string key && (key.ToUpper().StartsWith("ACS_", StringComparison.Ordinal) || key.ToUpper().StartsWith("ACS.", StringComparison.Ordinal)))
                                    {
                                        if (key.ToUpper().StartsWith("ACS.", StringComparison.Ordinal))
                                        {
                                            key = key.Replace('.', '_');
                                        }

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
                                            _ = projectIds.Add(key);
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
                    {
                        ConsoleLogger.LogWarning("No sheets found for the filter condition in monthly report " + report + ".");
                    }
                }
                catch (Exception ex)
                {
                    ConsoleLogger.LogErrorAndExit("Error on reading monthly report " + report + ": " + ex.Message);
                }
            }
            return new MonthlyReportData()
            {
                EmployeesData = employeeDataList,
                ProjectIds = projectIds
            };
        }

        public static PtrData ReadPtr(List<string> reports, int bookingMonthCol, object[] bookingMonths, object[] effortCols, int projectIdCol, string sheetName)
        {
            PtrData ptrData = new();
            _ptrBookingMonthCol = bookingMonthCol;
            bool flag = bookingMonths.Any();
            if (!flag && ConvertInputBookingMonths(bookingMonths))
            {
                ConsoleLogger.LogLine();
            }
            else
            {
                ConsoleLogger.LogErrorAndExit($"Error converting Ptr booking months: {bookingMonths}");
            }
            try
            {
                Dictionary<string, TimeSpan> projectEfforts = new();
                HashSet<string> projectIds = new();
                foreach (string ptrFile in reports)
                {
                    ConsoleLogger.LogInfo("Reading " + new FileInfo(ptrFile).Name);
                    using FileStream fileStream = File.Open(ptrFile, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTable? table = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables[sheetName];
                    if (table == null)
                    {
                        ConsoleLogger.LogErrorAndExit("Error on reading Ptr: no sheet found for name " + sheetName);
                    }
                    else
                    {
                        List<DataRow> list = table.AsEnumerable().Where(r => r[bookingMonthCol - 1] != DBNull.Value).Skip(1).ToList();
                        List<DataRow> rows = flag ? list : list.Where(IsRowOfBookingMonth).ToList();
                        projectIdCol--;
                        if (rows.Any())
                        {
                            rows.Select((r, i) => new { item = r, index = i }).ToList()
                                .ForEach(obj =>
                                {
                                    DataRow row = obj.item;
                                    int rowIndex = obj.index;
                                    if (!((string)row[projectIdCol]).ToUpper().StartsWith("ACS", StringComparison.Ordinal))
                                    {
                                        return;
                                    }

                                    string key = ((string)row[projectIdCol]).Replace('.', '_');
                                    TimeSpan totalEffort = TimeSpan.Zero;
                                    Array.ForEach(effortCols, ef =>
                                    {
                                        int effortCol = (int)ef - 1;
                                        if (row[effortCol] == DBNull.Value)
                                        {
                                            totalEffort += TimeSpan.Zero;
                                        }
                                        else if (row[effortCol] is double doubleValue)
                                        {
                                            totalEffort += new TimeSpan((int)doubleValue, 0, 0);
                                        }
                                        else if (row[effortCol] is TimeSpan timeSpan)
                                        {
                                            totalEffort += timeSpan;
                                        }
                                        else if (row[effortCol] is DateTime dateTime)
                                        {
                                            totalEffort += new TimeSpan(dateTime.Hour, dateTime.Minute, dateTime.Second);
                                        }
                                        else
                                        {
                                            ConsoleLogger.LogErrorAndExit($"Unknown format of the effort value at column: {ef} at row: {rowIndex}");
                                        }
                                    });
                                    if (projectEfforts.ContainsKey(key))
                                    {
                                        projectEfforts[key] += totalEffort;
                                    }
                                    else
                                    {
                                        projectEfforts[key] = totalEffort;
                                    }

                                    _ = projectIds.Add(key);
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

        public static UserSettings? GetUserSettings()
        {
            UserSettings? userSettings = default;
            try
            {
                userSettings = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText("userSettings.json"));
                if (userSettings is not null)
                {
                    ConsoleLogger.LogInfo("User settings:", 1);
                    for (int i = 0; i < userSettings.Actions.Length; i++)
                    {
                        ConsoleLogger.LogChar('-', 100);
                        ConsoleLogger.LogLine(1);
                        PrintAction(userSettings.Actions[i]);
                        if (i == userSettings.Actions.Length - 1)
                        {
                            ConsoleLogger.LogChar('-', 100);
                            ConsoleLogger.LogLine(1);
                        }
                    }
                }
                else
                {
                    ConsoleLogger.LogErrorAndExit("Error on reading user settings: null userSettings value");
                }
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on reading user settings: " + ex.Message);
            }
            return userSettings;
        }

        private static void PrintAction(Models.Action action)
        {
            string convertObjectToString(object obj)
            {
                return obj?.ToString() ?? string.Empty;
            }
            string[] convertObjectsToStrings(object[] objects) => Array.ConvertAll(objects, convertObjectToString);
            ConsoleLogger.LogSameLine("Name: "); ConsoleLogger.LogDataSameLine(action.Name, 1);
            ConsoleLogger.LogSameLine("Run: "); ConsoleLogger.LogDataSameLine(action.Run.ToString(), 1);
            ConsoleLogger.LogSameLine("InputFolder: "); ConsoleLogger.LogDataSameLine(action.InputFolder, 1);

            if (action.MonthlyReportMonths is not null)
            {
                ConsoleLogger.LogSameLine("MonthlyReportMonths: ");
                ConsoleLogger.LogDataSameLine(string.Join(",", convertObjectsToStrings(action.MonthlyReportMonths)), 1);
            }
            if (action.MonthlyReportIdCol is not null)
            {
                ConsoleLogger.LogSameLine("MonthlyReportIdCol: ");
                ConsoleLogger.LogDataSameLine(action.MonthlyReportIdCol?.ToString() ?? string.Empty, 1);
            }
            if (action.PtrSheetName is not null)
            {
                ConsoleLogger.LogSameLine("PtrSheetName: ");
                ConsoleLogger.LogDataSameLine(action.PtrSheetName, 1);
            }
            if (action.PtrProjectIdCol is not null)
            {
                ConsoleLogger.LogSameLine("PtrProjectIdCol: ");
                ConsoleLogger.LogDataSameLine(action.PtrProjectIdCol.ToString() ?? string.Empty, 1);
            }
            if (action.PtrBookingMonthCol is not null)
            {
                ConsoleLogger.LogSameLine("PtrBookingMonthCol: ");
                ConsoleLogger.LogDataSameLine(action.PtrBookingMonthCol.ToString() ?? string.Empty, 1);
            }
            if (action.PtrBookingMonths is not null)
            {
                ConsoleLogger.LogSameLine("PtrBookingMonths: ");
                ConsoleLogger.LogDataSameLine(string.Join(",", convertObjectsToStrings(action.PtrBookingMonths)), 1);
            }
            if (action.PtrEffortCols is not null)
            {
                ConsoleLogger.LogSameLine("PtrEffortCols: ");
                ConsoleLogger.LogDataSameLine(string.Join(",", convertObjectsToStrings(action.PtrEffortCols)), 1);
            }
            if (action.FinancialYear is not null)
            {
                ConsoleLogger.LogSameLine("FinancialYear: ");
                ConsoleLogger.LogDataSameLine(action.FinancialYear, 1);
            }
        }

        private static bool ConvertInputBookingMonths(object[] inputBookingMonths)
        {
            Array.ForEach(inputBookingMonths, ibm =>
            {
                JsonElement jsonElement = (JsonElement)ibm;
                switch (jsonElement.ValueKind)
                {
                    case JsonValueKind.String:
                        string? str = jsonElement.GetString();
                        if (string.IsNullOrEmpty(str))
                        {
                            break;
                        }

                        if (str.Contains('|'))
                        {
                            string[] splits = str.Split('|');
                            string numMonth = splits[0].Trim();
                            if (int.TryParse(numMonth, out int month))
                            {
                                BookingMonths.Add(month);
                            }

                            BookingMonths.AddRange(splits[1..]);
                            break;
                        }
                        BookingMonths.Add(str.Trim());
                        break;

                    case JsonValueKind.Number:
                        if (!DataService.Months.Contains(jsonElement.GetInt32()))
                        {
                            break;
                        }

                        BookingMonths.Add(jsonElement.GetInt32());
                        break;

                    default:
                        break;
                }
            });
            return true;
        }

        private static bool IsRowOfBookingMonth(DataRow dataRow)
        {
            return dataRow[_ptrBookingMonthCol - 1] switch
            {
                double num => BookingMonths.Contains((int)num),
                string str => BookingMonths.Contains(str),
                _ => false,
            };
        }

        public static EmployeePunchData ReadPunchMovementReport()
        {
            EmployeePunchData employeePunchData = new EmployeePunchData();

            return employeePunchData;
        }
    }
}