using ExcelDataReader;
using Models;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
using Utilities;

namespace Services
{
    public static class ExportService
    {
        public static void ExportConsolidateData(List<ConsolidatedData> consolidatedDataList, PtrData ptrData, MonthlyReportData monthlyReportData, string time, string exportFolder)
        {
            ConsoleLogger.LogInfo("Exporting consolidated data:", 1);
            try
            {
                string tableName = $"ConsolidatedReport_{time}";
                string exportPath = $"{exportFolder}\\{tableName}.xls";
                DataTable dataTable = new(tableName);
                DataColumn column1 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Project Id",
                    Caption = "Project Id",
                    ReadOnly = false
                };
                dataTable.Columns.Add(column1);
                DataColumn column2 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Total Effort as per PTR",
                    Caption = "Total Effort as per PTR",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column2);
                DataColumn column3 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Total Actual Effort as per Monthly report",
                    Caption = "Total Actual Effort as per Monthly report",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column3);
                List<EmployeeActualEffort> list = consolidatedDataList.SelectMany(data => (IEnumerable<EmployeeActualEffort>)data.EmployeeActualEffort).ToList();
                List<string> employeeNames = list.Select(eae => $"{eae.Name}({eae.Id})").Distinct().ToList();
                foreach (string str in employeeNames)
                {
                    DataColumn column4 = new()
                    {
                        DataType = typeof(string),
                        ColumnName = str,
                        Caption = str,
                        ReadOnly = false
                    };
                    dataTable.Columns.Add(column4);
                }
                TimeSpan TotalEffort = TimeSpan.Zero;
                TimeSpan TotalActualEffort = TimeSpan.Zero;
                DataRow consDtRow;
                foreach (ConsolidatedData consolidatedData in consolidatedDataList)
                {
                    consDtRow = dataTable.NewRow();
                    consDtRow["Project Id"] = consolidatedData.ProjectId;
                    consDtRow["Total Effort as per PTR"] = $"{(int)consolidatedData.TotalEffort.TotalHours}:{consolidatedData.TotalEffort.Minutes}";
                    TotalEffort += consolidatedData.TotalEffort;
                    TimeSpan totalActualEffort = new();
                    list.Where(eae => eae.ProjectId.Equals(consolidatedData.ProjectId, StringComparison.Ordinal)).ToList().ForEach(eae =>
                    {
                        totalActualEffort += eae.ActualEffort;
                        consDtRow[$"{eae.Name}({eae.Id})"] = $"{(int)eae.ActualEffort.TotalHours}:{eae.ActualEffort.Minutes}";
                    });
                    consDtRow["Total Actual Effort as per Monthly report"] = $"{(int)totalActualEffort.TotalHours}:{totalActualEffort.Minutes}";
                    TotalActualEffort += totalActualEffort;
                    dataTable.Rows.Add(consDtRow);
                }
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Total Hours";
                consDtRow["Total Effort as per PTR"] = $"{(int)TotalEffort.TotalHours}:{TotalEffort.Minutes}";
                consDtRow["Total Actual Effort as per Monthly report"] = $"{(int)TotalActualEffort.TotalHours}:{TotalActualEffort.Minutes}";
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                    .ForEach(ed =>
                    {
                        TimeSpan timeSpan2 = ed.ProjectData
                        .Join(ptrData.ProjectEfforts, projectTime => projectTime.Key, projectEffort => projectEffort.Key, (projectTime, projectEffort) => projectTime.Value)
                        .AsEnumerable()
                        .Aggregate(new TimeSpan(), (current, item) => current + item);
                        consDtRow[$"{ed.Name}({ed.Id})"] = $"{(int)timeSpan2.TotalHours}:{timeSpan2.Minutes}";
                    });
                dataTable.Rows.Add(consDtRow);
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Actual Available Hours";
                consDtRow["Total Effort as per PTR"] = string.Empty ;
                TimeSpan totalActualAvailableHours = new();
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                   .ForEach(ed =>
                   {
                       TimeSpan timeSpan2 = ed.ActualAvailableHours;
                       consDtRow[$"{ed.Name}({ed.Id})"] = $"{(int)timeSpan2.TotalHours}:{timeSpan2.Minutes}";
                       totalActualAvailableHours += timeSpan2;
                   });
                consDtRow["Total Actual Effort as per Monthly report"] = $"{(int)totalActualAvailableHours.TotalHours}:{totalActualAvailableHours.Minutes}";
                dataTable.Rows.Add(consDtRow);
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Total Leaves availed by team in Days";
                int totalLeaves = 0;
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                   .ForEach(ed =>
                   {
                       consDtRow[$"{ed.Name}({ed.Id})"] = ed.TotalLeaves;
                       totalLeaves += ed.TotalLeaves;
                   });
                consDtRow["Total Actual Effort as per Monthly report"] = totalLeaves;
                consDtRow["Total Effort as per PTR"] = (totalLeaves * 9).ToString();
                dataTable.Rows.Add(consDtRow);
                WriteExcel(dataTable, exportPath, tableName);
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on exporting data: " + ex.Message);
            }
        }

        public static void ExportLeaveReport(List<string> monthlyReports, string financialYear, string exportFolder)
        {
            ConsoleLogger.LogInfo("Generating Leave Report for FY" + financialYear + ":", 1);
            List<string> sheetNames = DataService.GetFyMonths(financialYear);
            List<LeaveReportData> leaveReportDataList = new();
            bool hasReadErrors = false;
            foreach (string monthlyReport in monthlyReports)
            {
                ConsoleLogger.Log("Reading " + new FileInfo(monthlyReport).Name);
                LeaveReportData leaveReportData;
                using (FileStream fileStream = File.Open(monthlyReport, FileMode.Open, FileAccess.Read))
                {
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> sheets = tables.Cast<DataTable>().Where(dataTable => sheetNames.Contains(dataTable.TableName.Trim())).ToList();
                    int? totalLeaves = new int?(0);
                    Dictionary<string, int?> leaves = new();
                    string employeeName = string.Empty;
                    int employeeId = default;
                    if (sheets.Count > 0)
                    {
                        for (int i = 0; i < sheets.Count; i++)
                        {
                            if (sheets[i].Rows[3][2] is DBNull || sheets[i].Rows[4][2] is DBNull)
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(((string)sheets[i].Rows[3][2]).Trim()))
                            {
                                throw new ArgumentException("Employee name is empty or has an invalid format in the sheet " + sheets[i].TableName + ": Check row " + 4 + " at column " + 3);
                            }
                            employeeName = (string)sheets[i].Rows[3][2];
                            employeeName = employeeName.Trim();
                            if (!int.TryParse(sheets[i].Rows[4][2].ToString(), out employeeId))
                            {
                                throw new ArgumentException("Employee Id is empty or has an invalid format in the sheet " + sheets[i].TableName + ": Check row " + 5 + " at column " + 3);
                            }
                            break;
                        }
                        ConsoleLogger.LogSameLine("Reading Sheet: ");
                        sheetNames.ForEach(sheetName =>
                        {
                            if (sheets.Exists(sh => sh.TableName == sheetName))
                            {
                                DataTable dataTable = sheets.First(sh => sh.TableName == sheetName);
                                ConsoleLogger.LogDataSameLine(dataTable.TableName + ", ");
                                try
                                {
                                    DataRowCollection rows = dataTable.Rows;
                                    int lastColumnIndex = rows[14].ItemArray.Length - 1;
                                    if (int.TryParse(rows[14][lastColumnIndex].ToString(), out int monthlyLeave))
                                    {
                                        leaves[dataTable.TableName] = new int?(monthlyLeave);
                                        totalLeaves = totalLeaves.HasValue ? new int?(totalLeaves.GetValueOrDefault() + monthlyLeave) : new int?();
                                    }
                                    else
                                    {
                                        throw new FormatException($"Invalid format at column: {lastColumnIndex + 1} at row: 15in sheet: {dataTable.TableName}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    hasReadErrors = true;
                                    ConsoleLogger.LogErrorAndExit("Error on generating leave report for " + monthlyReport + ": " + ex.Message);
                                }
                            }
                            else
                            {
                                leaves[sheetName] = new int?();
                                totalLeaves = new int?();
                            }
                        });
                        ConsoleLogger.LogLine();
                    }
                    leaveReportData = new LeaveReportData()
                    {
                        EmployeeId = employeeId.ToString(CultureInfo.InvariantCulture),
                        Name = employeeName,
                        Leaves = leaves,
                        TotalLeaves = totalLeaves
                    };
                }
                leaveReportDataList.Add(leaveReportData);
            }
            if (!hasReadErrors)
            {
                string tableName = $"LeaveReport-FY{financialYear}";
                string exportPath = $"{exportFolder}\\{tableName}.xls";
                DataTable dataTable = new(tableName);
                DataColumn column1 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Employee Id",
                    Caption = "Employee Id",
                    ReadOnly = false
                };
                dataTable.Columns.Add(column1);
                DataColumn column2 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Employee Name",
                    Caption = "Employee Name",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column2);
                foreach (string str in sheetNames)
                {
                    DataColumn column3 = new()
                    {
                        DataType = typeof(string),
                        ColumnName = str,
                        Caption = str,
                        ReadOnly = false,
                        Unique = false
                    };
                    dataTable.Columns.Add(column3);
                }
                DataColumn column4 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Total Leave Days",
                    Caption = "Total Leave Days",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column4);
                foreach (LeaveReportData leaveReportData in leaveReportDataList)
                {
                    DataRow row = dataTable.NewRow();
                    row["Employee Id"] = leaveReportData.EmployeeId;
                    row["Employee Name"] = leaveReportData.Name;
                    foreach (KeyValuePair<string, int?> leaf in leaveReportData.Leaves)
                        row[leaf.Key] = leaf.Value.HasValue ? leaf.Value.ToString() : "NA";
                    row["Total Leave Days"] = leaveReportData.TotalLeaves.HasValue ? leaveReportData.TotalLeaves.ToString() : "NA";
                    dataTable.Rows.Add(row);
                }
                WriteExcel(dataTable, exportPath, tableName);
            }
            else
            {
                ConsoleLogger.LogErrorAndExit("Process stopped due to errors.");
            }
        }

        public static void ExportMonthlyReportInter(MonthlyReportData monthlyReportData, string time, string exportFolder)
        {
            ConsoleLogger.LogInfo("Exporting monthly report inter:", 2);
            string tableName = $"MonthlyReport_Inter_{time}";
            string exportPath = $"{exportFolder}\\{tableName}.xls";
            DataTable monthlyTable = new(tableName);
            DataColumn column1 = new()
            {
                DataType = typeof(string),
                ColumnName = "Name(Id)",
                Caption = "Name(Id)",
                ReadOnly = false
            };
            monthlyTable.Columns.Add(column1);
            DataColumn column2 = new()
            {
                DataType = typeof(string),
                ColumnName = "Project Id",
                Caption = "Project Id",
                ReadOnly = false
            };
            monthlyTable.Columns.Add(column2);
            DataColumn column3 = new()
            {
                DataType = typeof(string),
                ColumnName = "Actual Effort",
                Caption = "Actual Effort",
                ReadOnly = false,
                Unique = false
            };
            monthlyTable.Columns.Add(column3);
            foreach (EmployeeData employeeData1 in monthlyReportData.EmployeesData.Where<EmployeeData>(employeeData => employeeData.ProjectData.Count > 0))
            {
                EmployeeData employeeData = employeeData1;
                employeeData.ProjectData.ToList().ForEach(pt =>
                {
                    DataRow monthlyDtRow = monthlyTable.NewRow();
                    monthlyDtRow["Name(Id)"] = $"{employeeData.Name}({employeeData.Id})";
                    monthlyDtRow["Project Id"] = pt.Key;
                    monthlyDtRow["Actual Effort"] = $"{(int)pt.Value.TotalHours}:{pt.Value.Minutes}";
                    monthlyTable.Rows.Add(monthlyDtRow);
                });
            }
            WriteExcel(monthlyTable, exportPath, tableName);
        }

        public static void ExportPtrInter(PtrData ptrData, string time, string exportFolder)
        {
            ConsoleLogger.LogInfo("Exporting ptr inter data:", 1);
            string tableName = $"PTR_Inter_{time}";
            string exportPath = $"{exportFolder}\\{tableName}.xls";
            DataTable dataTable = new(tableName);
            DataColumn column1 = new()
            {
                DataType = typeof(string),
                ColumnName = "Project Id",
                Caption = "Project Id",
                ReadOnly = false
            };
            dataTable.Columns.Add(column1);
            DataColumn column2 = new()
            {
                DataType = typeof(string),
                ColumnName = "Total Effort",
                Caption = "Total Effort",
                ReadOnly = false,
                Unique = false
            };
            dataTable.Columns.Add(column2);
            foreach (KeyValuePair<string, TimeSpan> projectEffort in ptrData.ProjectEfforts)
            {
                DataRow row = dataTable.NewRow();
                row["Project Id"] = projectEffort.Key;
                row["Total Effort"] = $"{(int)projectEffort.Value.TotalHours}:{projectEffort.Value.Minutes}";
                dataTable.Rows.Add(row);
            }
            WriteExcel(dataTable, exportPath, tableName);
        }

        private static void WriteExcel(DataTable dataTable, string filePath, string reportName = "data")
        {
            if (dataTable is not null && !string.IsNullOrEmpty(filePath))
            {
                if (!Directory.Exists(Path.GetDirectoryName(filePath)))
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
                using (StreamWriter streamWriter1 = new(filePath))
                {
                    for (int index = 0; index < dataTable.Columns.Count; ++index)
                    {
                        StreamWriter streamWriter2 = streamWriter1;
                        DefaultInterpolatedStringHandler interpolatedStringHandler = new(1, 1);
                        interpolatedStringHandler.AppendFormatted(dataTable.Columns[index]);
                        interpolatedStringHandler.AppendLiteral("\t");
                        string stringAndClear = interpolatedStringHandler.ToStringAndClear();
                        streamWriter2.Write(stringAndClear);
                    }
                    streamWriter1.WriteLine();
                    for (int index = 0; index < dataTable.Rows.Count; ++index)
                    {
                        for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; ++columnIndex)
                            streamWriter1.Write(Convert.ToString(dataTable.Rows[index][columnIndex], CultureInfo.InvariantCulture) + "\t");
                        streamWriter1.WriteLine();
                    }
                }
                ConsoleLogger.LogSameLine($"Exported {reportName}: ", 0); ConsoleLogger.LogDataSameLine(filePath);
            }
        }
    }
}