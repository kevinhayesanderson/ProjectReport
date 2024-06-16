using ExcelDataReader;
using Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Schema;
using System.Data;
using Utilities;

namespace Services
{
    public class ReadService(ILogger logger)
    {
        private List<object>? _bookingMonths;

        private int _ptrBookingMonthCol;

        public UserSettings? GetUserSettings()
        {
            UserSettings? userSettings = default;
            try
            {
                var userSettingsInfo = new UserSettingsInfo();
                JSchema schema = JSchema.Parse(userSettingsInfo.SchemaValue);
                using JsonTextReader reader = new(File.OpenText(userSettingsInfo.UserSettingsFileName));

                JSchemaValidatingReader validatingReader = new(reader)
                {
                    Schema = schema
                };

                List<string> messages = [];
                validatingReader.ValidationEventHandler += (o, a) => messages.Add(a.Message);

                JsonSerializer serializer = new();
                userSettings = serializer.Deserialize<UserSettings>(validatingReader);

                if (messages.Count > 0)
                {
                    logger.LogError("Validation error on reading user settings:", 1);
                    messages.ForEach(message => logger.LogError(message));
                    logger.LogError("Refer documentation to solve above error");
                    return null;
                }

                if (userSettings is not null)
                {
                    logger.LogInfo("Active actions:", 1);
                    var activeActions = userSettings.Actions.Where(action => action.Run).ToArray();
                    for (int i = 0; i < activeActions.Length; i++)
                    {
                        logger.LogChar('-', 100);
                        logger.LogLine(1);
                        PrintAction(activeActions[i]);
                        if (i == activeActions.Length - 1)
                        {
                            logger.LogChar('-', 100);
                            logger.LogLine(1);
                        }
                    }
                }
                else
                {
                    logger.LogError("Error on reading user settings: null userSettings value");
                    return null;
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Error on reading user settings: {ex}");
                throw;
            }
            return userSettings;
        }

        public MonthlyReportData ReadMonthlyReports(List<string> reports, object[] months, int idCol)
        {
            List<EmployeeData> employeeDataList = [];
            HashSet<string> projectIds = [];
            foreach (string report in reports)
            {
                try
                {
                    logger.LogInfo($"Reading {new FileInfo(report).Name}", 1);
                    using FileStream fileStream = File.Open(report, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> dataTableList = months.Length != 0
                        ? tables.Cast<DataTable>().Where(dataTable => months.Contains(dataTable.TableName.Trim())).ToList()
                        : tables.Cast<DataTable>().Select(dataTable => dataTable).ToList();
                    if (dataTableList.Count > 0)
                    {
                        if (string.IsNullOrEmpty(((string)dataTableList[0].Rows[3][2]).Trim()))
                        {
                            throw new ArgumentException($"Employee name is empty or has an invalid format in the sheet {dataTableList[0].TableName}: Check row {4} at column {3}");
                        }
                        string employeeName = (string)dataTableList[0].Rows[3][2];
                        employeeName = employeeName.Trim();
                        if (!int.TryParse(dataTableList[0].Rows[4][2].ToString(), out int employeeId))
                        {
                            throw new ArgumentException($"Employee Id is empty or has an invalid format in the sheet {dataTableList[0].TableName}: Check row {5} at column {3}");
                        }
                        Dictionary<string, TimeSpan> projectData = [];
                        logger.LogSameLine("Reading Sheet: ");
                        TimeSpan ActualAvailableHours = TimeSpan.Zero;
                        int TotalLeaves = new();
                        foreach (DataTable dataTable in dataTableList)
                        {
                            try
                            {
                                logger.LogDataSameLine(dataTable.TableName + ", ");
                                DataRowCollection rows = dataTable.Rows;
                                int lastColumnIndex;
                                int actualAvailableHoursRowIndex = 13;
                                lastColumnIndex = rows[actualAvailableHoursRowIndex].ItemArray.Length - 1;
                                if (TimeSpan.TryParse(rows[actualAvailableHoursRowIndex][lastColumnIndex].ToString(), out TimeSpan actualAvailableHours))
                                {
                                    if (actualAvailableHours.Equals(new TimeSpan()))
                                    {
                                        logger.LogWarning($"Check for potential empty data at column: {lastColumnIndex + 1} at row: {actualAvailableHoursRowIndex + 1} in sheet {dataTable.TableName}", 1);
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
                                        logger.LogWarning($"Check for potential empty data at column: {lastColumnIndex + 1} at row: {totalLeavesRowIndex + 1} in sheet {dataTable.TableName}", 1);
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
                                                logger.LogWarning($"Check for potential empty data at column: {lastColumnIndex + 1} at row: {rowIndex + 1} in sheet {dataTable.TableName}", 1);
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
                                logger.LogError($"Error on reading sheet {dataTable.TableName}: {ex}");
                                throw;
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
                        logger.LogWarning($"No sheets found for the filter condition in monthly report {report}.");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError($"Error on reading monthly report {report}: {ex}");
                    throw;
                }
            }
            return new MonthlyReportData()
            {
                EmployeesData = employeeDataList,
                ProjectIds = projectIds
            };
        }

        public MusterOptionsDatas ReadMusterOptions(List<string> musterOptionsReports)
        {
            MusterOptionsDatas musterOptionsDatas = new MusterOptionsDatas();
            foreach (var musterOptionsReport in musterOptionsReports)
            {
                try
                {
                    logger.LogInfo($"Reading {new FileInfo(musterOptionsReport).Name}", 1);
                    using FileStream fileStream = File.Open(musterOptionsReport, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> dataTableList = tables.Cast<DataTable>().Select(dataTable => dataTable).ToList();
                    if (dataTableList.Count > 0)
                    {
                        int i = 0;
                        logger.LogSameLine("Reading Sheet: ");
                        int eCodeColumn = -1;
                        int nameColumn = -1;
                        int designationColumn = -1;
                        int firstDateColumn = -1;
                        int lastDateColumn = -1;
                        DateTime firstDate = default;
                        DateTime lastDate = default;
                        foreach (DataTable dataTable in dataTableList)
                        {
                            try
                            {
                                if (i == 0)
                                {
                                    var trimmedColumnNames = dataTable.Rows[3].ItemArray.Select(o => o?.ToString()?.Trim().ToLower()).ToArray();
                                    eCodeColumn = Array.IndexOf(trimmedColumnNames, "empcode");
                                    nameColumn = Array.IndexOf(trimmedColumnNames, "name");
                                    designationColumn = Array.IndexOf(trimmedColumnNames, "designation");
                                    firstDateColumn = Array.FindIndex(trimmedColumnNames, item => DateTime.TryParse(item, out firstDate));
                                    lastDateColumn = Array.FindLastIndex(trimmedColumnNames, item => DateTime.TryParse(item, out lastDate));
                                }
                                logger.LogDataSameLine(dataTable.TableName + ", ");
                                DataRowCollection rows = dataTable.Rows;

                                for (int j = 5; j < rows.Count; j += 6)
                                {
                                    var row = rows[j];
                                    var employeeId = Convert.ToUInt32(row[eCodeColumn]);
                                    if (row[nameColumn] is not String employeeName) continue;
                                    var designation = row[designationColumn] as String ?? string.Empty;
                                    var musterOptions = new List<MusterOption>();
                                    for (int k = firstDateColumn; k <= lastDateColumn; k++)
                                    {
                                        var date = (DateTime)rows[3][k];
                                        var inTime = DataService.ConvertToTimeOnly(rows[j + 2][k]);
                                        var outTime = DataService.ConvertToTimeOnly(rows[j + 3][k]);
                                        var musterOption = new MusterOption()
                                        {
                                            Date = date,
                                            InTime = inTime,
                                            OutTime = outTime
                                        };
                                        musterOptions.Add(musterOption);
                                    }

                                    if (musterOptionsDatas.Datas.TryGetValue(employeeId, out MusterOptionsData? musterOptionsData))
                                    {
                                        musterOptionsData.AddMusterOptions(musterOptions);
                                    }
                                    else
                                    {
                                        musterOptionsDatas.Datas[employeeId] = new MusterOptionsData()
                                        {
                                            Name = employeeName,
                                            Designation = designation,
                                            MusterOptions = musterOptions
                                        };
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                logger.LogError($"Error on reading sheet {dataTable.TableName}: {ex}");
                                throw;
                            }
                            i++;
                        }
                    }
                    else
                    {
                        logger.LogWarning($"No sheets found in muster options report {musterOptionsReport}.");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError($"Error on reading muster options report {musterOptionsReport}: {ex}");
                    throw;
                }
            }
            return musterOptionsDatas;
        }

        public PtrData ReadPtr(List<string> reports, int bookingMonthCol, object[] bookingMonths, object[] effortCols, int projectIdCol, string sheetName)
        {
            PtrData ptrData = new();
            _ptrBookingMonthCol = bookingMonthCol;
            try
            {
                if (bookingMonths.Length != 0 && ConvertInputBookingMonths(bookingMonths))
                {
                    logger.LogLine();
                }
                else
                {
                    string errMessage = $"Error converting Ptr booking months: {bookingMonths}";
                    logger.LogError(errMessage);
                    throw new FormatException(errMessage);
                }

                Dictionary<string, TimeSpan> projectEfforts = [];
                HashSet<string> projectIds = [];
                foreach (string ptrFile in reports)
                {
                    logger.LogInfo($"Reading {new FileInfo(ptrFile).Name}");
                    using FileStream fileStream = File.Open(ptrFile, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTable? table = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables[sheetName];
                    if (table == null)
                    {
                        string errMessage = $"Error on reading Ptr: no sheet found for name {sheetName}";
                        logger.LogError(errMessage);
                        throw new FormatException(errMessage);
                    }
                    else
                    {
                        List<DataRow> list = table.AsEnumerable().Where(r => r[bookingMonthCol - 1] != DBNull.Value).Skip(1).ToList();
                        List<DataRow> rows = bookingMonths.Length != 0 ? list.Where(IsRowOfBookingMonth).ToList() : list;
                        projectIdCol--;
                        if (rows.Count != 0)
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
                                        int effortCol = Convert.ToInt32(ef) - 1;
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
                                            string errMessage = $"Unknown format of the effort value at column: {ef} at row: {rowIndex}";
                                            logger.LogError(errMessage);
                                            throw new FormatException(errMessage);
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
                logger.LogError($"Error on reading Ptr: {ex}");
                throw;
            }
            return ptrData;
        }

        public PunchMovementData ReadPunchMovementReports(List<string> reports)
        {
            List<EmployeePunchData> employeePunchDatas = [];
            foreach (string report in reports)
            {
                try
                {
                    logger.LogInfo($"Reading {new FileInfo(report).Name}", 1);
                    using FileStream fileStream = File.Open(report, FileMode.Open, FileAccess.Read);
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> dataTableList = tables.Cast<DataTable>().Select(dataTable => dataTable).ToList();
                    if (dataTableList.Count > 0)
                    {
                        int i = 0;
                        logger.LogSameLine("Reading Sheet: ");
                        int eCodeColumn = -1;
                        int nameColumn = -1;
                        int dateColumn = -1;
                        int firstInOutColumn = -1;
                        int lastInOutColumn = -1;
                        foreach (DataTable dataTable in dataTableList)
                        {
                            try
                            {
                                if (i == 0)
                                {
                                    var trimmedColumnNames = dataTable.Rows[1].ItemArray.Select(o => o?.ToString()?.Trim().ToLower()).ToArray();
                                    eCodeColumn = Array.IndexOf(trimmedColumnNames, "ecode");
                                    nameColumn = Array.IndexOf(trimmedColumnNames, "name");
                                    dateColumn = Array.IndexOf(trimmedColumnNames, "date");
                                    firstInOutColumn = Array.FindIndex(trimmedColumnNames, item => item is "in" or "out");
                                    lastInOutColumn = Array.FindLastIndex(trimmedColumnNames, item => item is "in" or "out");
                                }
                                logger.LogDataSameLine(dataTable.TableName + ", ");
                                DataRowCollection rows = dataTable.Rows;
                                int j = 0;
                                EmployeePunchData employeePunchData = new();
                                foreach (DataRow row in rows)
                                {
                                    if (j == 0)
                                    {
                                        j++;
                                        continue;
                                    }
                                    if (DateTime.TryParse(row.ItemArray[dateColumn]?.ToString(), out DateTime date))
                                    {
                                        var name = row.ItemArray[nameColumn]?.ToString();
                                        if (!string.IsNullOrEmpty(name) && int.TryParse(row.ItemArray[eCodeColumn]?.ToString(), out int id))
                                        {
                                            employeePunchData = new EmployeePunchData()
                                            {
                                                Name = name,
                                                Id = id,
                                                PunchDatas = []
                                            };
                                            employeePunchDatas.Add(employeePunchData);
                                        }
                                        employeePunchData.PunchDatas.Add(
                                            new PunchData()
                                            {
                                                Date = date,
                                                Punches = row.ItemArray[firstInOutColumn..(lastInOutColumn + 1)]
                                                            .Where(item => item != null && DateTime.TryParse(item.ToString(), out _))
                                                            .Select(item => TimeOnly.FromDateTime(DateTime.Parse(item?.ToString()!)))
                                                            .ToList()
                                            });
                                    }
                                    j++;
                                }
                            }
                            catch (Exception ex)
                            {
                                logger.LogError($"Error on reading sheet {dataTable.TableName}: {ex}");
                                throw;
                            }
                            i++;
                        }
                    }
                    else
                    {
                        logger.LogWarning($"No sheets found in punch movement report {report}.");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError($"Error on reading punch movement report {report}: {ex}");
                    throw;
                }
            }
            return new PunchMovementData([.. employeePunchDatas]);
        }

        private bool ConvertInputBookingMonths(object[] inputBookingMonths)
        {
            _bookingMonths = [];
            Array.ForEach(inputBookingMonths, ibm =>
            {
                switch (ibm)
                {
                    case string stringValue:
                        string? str = stringValue;
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
                                _bookingMonths.Add(month);
                            }

                            _bookingMonths.AddRange(splits[1..]);
                            break;
                        }
                        _bookingMonths.Add(str.Trim());
                        break;

                    case long number:
                        if (!DataService.Months.Contains(number))
                        {
                            break;
                        }

                        _bookingMonths.Add(number);
                        break;

                    default:
                        break;
                }
            });
            return true;
        }

        private bool IsRowOfBookingMonth(DataRow dataRow)
        {
            bool res = false;
            if (_bookingMonths != null)
            {
                res = dataRow[_ptrBookingMonthCol - 1] switch
                {
                    double num => _bookingMonths.Contains((int)num) || _bookingMonths.Contains((long)num),
                    string str => _bookingMonths.Contains(str),
                    _ => false,
                };
            }
            return res;
        }

        private void PrintAction(Models.Action action)
        {
            string convertObjectToString(object obj)
            {
                return obj?.ToString() ?? string.Empty;
            }
            string[] convertObjectsToStrings(object[] objects) => Array.ConvertAll(objects, convertObjectToString);
            logger.LogSameLine("Name: "); logger.LogDataSameLine(action.Name, 1);
            logger.LogSameLine("Run: "); logger.LogDataSameLine(action.Run.ToString(), 1);
            logger.LogSameLine("InputFolder: "); logger.LogDataSameLine(action.InputFolder, 1);

            if (action.MonthlyReportMonths is not [])
            {
                logger.LogSameLine("MonthlyReportMonths: ");
                logger.LogDataSameLine(string.Join(",", convertObjectsToStrings(action.MonthlyReportMonths)), 1);
            }
            if (action.MonthlyReportIdCol is not -1)
            {
                logger.LogSameLine("MonthlyReportIdCol: ");
                logger.LogDataSameLine(action.MonthlyReportIdCol.ToString() ?? string.Empty, 1);
            }
            if (!string.IsNullOrEmpty(action.PtrSheetName))
            {
                logger.LogSameLine("PtrSheetName: ");
                logger.LogDataSameLine(action.PtrSheetName, 1);
            }
            if (action.PtrProjectIdCol is not -1)
            {
                logger.LogSameLine("PtrProjectIdCol: ");
                logger.LogDataSameLine(action.PtrProjectIdCol.ToString() ?? string.Empty, 1);
            }
            if (action.PtrBookingMonthCol is not -1)
            {
                logger.LogSameLine("PtrBookingMonthCol: ");
                logger.LogDataSameLine(action.PtrBookingMonthCol.ToString() ?? string.Empty, 1);
            }
            if (action.PtrBookingMonths is not [])
            {
                logger.LogSameLine("PtrBookingMonths: ");
                logger.LogDataSameLine(string.Join(",", convertObjectsToStrings(action.PtrBookingMonths)), 1);
            }
            if (action.PtrEffortCols is not [])
            {
                logger.LogSameLine("PtrEffortCols: ");
                logger.LogDataSameLine(string.Join(",", convertObjectsToStrings(action.PtrEffortCols)), 1);
            }
            if (!string.IsNullOrEmpty(action.FinancialYear))
            {
                logger.LogSameLine("FinancialYear: ");
                logger.LogDataSameLine(action.FinancialYear, 1);
            }
            if (!string.IsNullOrEmpty(action.CutOff))
            {
                logger.LogSameLine("Cut-off: ");
                logger.LogDataSameLine(action.CutOff, 1);
            }
        }
    }
}