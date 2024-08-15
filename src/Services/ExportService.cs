using ExcelDataReader;
using Models;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
using Utilities;
using static Models.Constants;

namespace Services
{
    public class ExportService(ILogger logger)
    {
        public bool ExportConsolidateData(in List<ConsolidatedData> consolidatedDataList, ref MonthlyReportData monthlyReportData, in string time, in string exportFolder)
        {
            bool res = true;
            logger.LogInfo("Exporting Consolidated Report:", 1);
            try
            {
                string reportName = "ConsolidatedReport";
                string sheetName = $"{reportName}_Sheet1";
                string fileName = $"{reportName}_{time}.xls";
                string exportPath = $"{exportFolder}\\{fileName}";

                DataTable dataTable = new(sheetName);
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
                List<EmployeeActualEffort> EmployeeActualEfforts = consolidatedDataList.SelectMany(data => data.EmployeeActualEffort).ToList();
                List<string> employeeNames = EmployeeActualEfforts.Select(eae => $"{eae.Name}({eae.Id})").Distinct().Order().ToList();
                foreach (string str in employeeNames)
                {
                    DataColumn employeeColumn = new()
                    {
                        DataType = typeof(string),
                        ColumnName = str,
                        Caption = str,
                        ReadOnly = false
                    };
                    dataTable.Columns.Add(employeeColumn);
                }
                TimeSpan TotalEffort = TimeSpan.Zero;
                TimeSpan TotalActualEffort = TimeSpan.Zero;
                DataRow consDtRow;
                foreach (ConsolidatedData consolidatedData in consolidatedDataList)
                {
                    consDtRow = dataTable.NewRow();
                    consDtRow["Project Id"] = consolidatedData.ProjectId;
                    consDtRow["Total Effort as per PTR"] = DataService.TohhmmFormatString(consolidatedData.TotalEffort);
                    TotalEffort += consolidatedData.TotalEffort;
                    TimeSpan totalActualEffort = TimeSpan.Zero;
                    EmployeeActualEfforts.Where(eae => eae.ProjectId.Equals(consolidatedData.ProjectId, StringComparison.Ordinal)).ToList().ForEach(eae =>
                    {
                        totalActualEffort += eae.ActualEffort;
                        consDtRow[$"{eae.Name}({eae.Id})"] = DataService.TohhmmFormatString(eae.ActualEffort);
                    });
                    consDtRow["Total Actual Effort as per Monthly report"] = DataService.TohhmmFormatString(totalActualEffort);
                    TotalActualEffort += totalActualEffort;
                    dataTable.Rows.Add(consDtRow);
                }
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Total Hours";
                consDtRow["Total Effort as per PTR"] = DataService.TohhmmFormatString(TotalEffort);
                consDtRow["Total Actual Effort as per Monthly report"] = DataService.TohhmmFormatString(TotalActualEffort);
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                    .ForEach(ed =>
                    {
                        consDtRow[$"{ed.Name}({ed.Id})"] = DataService.TohhmmFormatString(ed.TotalProjectHours);
                    });
                dataTable.Rows.Add(consDtRow);
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Actual Available Hours";
                consDtRow["Total Effort as per PTR"] = string.Empty;
                TimeSpan totalActualAvailableHours = TimeSpan.Zero;
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                   .ForEach(ed =>
                   {
                       consDtRow[$"{ed.Name}({ed.Id})"] = DataService.TohhmmFormatString(ed.ActualAvailableHours);
                       totalActualAvailableHours += ed.ActualAvailableHours;
                   });
                consDtRow["Total Actual Effort as per Monthly report"] = DataService.TohhmmFormatString(totalActualAvailableHours);
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
                WriteExcel(ref dataTable, exportPath, fileName);
            }
            catch (Exception ex)
            {
                logger.LogError($"Error on exporting Consolidated Report: {ex}");
                return false;
            }
            return res;
        }

        public bool ExportLeaveReport(in List<string> monthlyReports, in string financialYear, in string exportFolder)
        {
            bool res = true;
            logger.LogInfo("Generating Leave Report for FY" + financialYear + ":", 1);
            List<string> sheetNames = DataService.GetMMMYYYforFy(financialYear);
            List<LeaveReportData> leaveReportDataList = [];
            bool hasReadErrors = false;
            foreach (string monthlyReport in monthlyReports)
            {
                logger.Log("Reading " + new FileInfo(monthlyReport).Name);
                LeaveReportData leaveReportData;
                using (FileStream fileStream = File.Open(monthlyReport, FileMode.Open, FileAccess.Read))
                {
                    using IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null);
                    DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                    List<DataTable> sheets = tables.Cast<DataTable>().Where(dataTable => sheetNames.Contains(dataTable.TableName.Trim())).ToList();
                    (int? totalLeaves, Dictionary<string, int?> leaves, string employeeName, int employeeId, int employeeIdColumnIndex, int employeeIdRowIndex, int employeeNameColumnIndex, int employeeNameRowIndex)
                    = (new int?(0), [], string.Empty, default, MonthlyReport.EmployeeIdIndex.Column, MonthlyReport.EmployeeIdIndex.Row, MonthlyReport.EmployeeNameIndex.Column, MonthlyReport.EmployeeNameIndex.Row);
                    if (sheets.Count > 0)
                    {
                        for (int i = 0; i < sheets.Count; i++)
                        {
                            if (sheets[i].Rows[employeeNameRowIndex][employeeNameColumnIndex] is DBNull || sheets[i].Rows[employeeIdRowIndex][employeeIdColumnIndex] is DBNull)
                            {
                                continue;
                            }
                            if (string.IsNullOrEmpty(((string)sheets[i].Rows[employeeNameRowIndex][employeeNameColumnIndex]).Trim()))
                            {
                                throw new ArgumentException("Employee name is empty or has an invalid format in the sheet " + sheets[i].TableName + ": Check row " + (employeeNameRowIndex + 1) + " at column " + (employeeNameColumnIndex + 1));
                            }
                            employeeName = (string)sheets[i].Rows[employeeNameRowIndex][employeeNameColumnIndex];
                            employeeName = employeeName.Trim();
                            if (!int.TryParse(sheets[i].Rows[employeeIdRowIndex][employeeIdColumnIndex].ToString(), out employeeId))
                            {
                                throw new ArgumentException("Employee Id is empty or has an invalid format in the sheet " + sheets[i].TableName + ": Check row " + (employeeIdRowIndex + 1) + " at column " + (employeeIdColumnIndex + 1));
                            }
                            break;
                        }
                        logger.LogSameLine("Reading Sheet: ");
                        sheetNames.ForEach(sheetName =>
                        {
                            if (sheets.Exists(sh => sh.TableName == sheetName))
                            {
                                DataTable dataTable = sheets.First(sh => sh.TableName == sheetName);
                                logger.LogDataSameLine(dataTable.TableName + ", ");
                                try
                                {
                                    DataRowCollection rows = dataTable.Rows;
                                    int lastColumnIndex = rows[MonthlyReport.LeavesRowIndex.Row].ItemArray.Length - 1;
                                    if (int.TryParse(rows[MonthlyReport.LeavesRowIndex.Row][lastColumnIndex].ToString(), out int monthlyLeave))
                                    {
                                        leaves[dataTable.TableName] = new int?(monthlyLeave);
                                        totalLeaves = totalLeaves.HasValue ? new int?(totalLeaves.GetValueOrDefault() + monthlyLeave) : new int?();
                                    }
                                    else
                                    {
                                        throw new FormatException($"Invalid format at column: {lastColumnIndex + 1} at row: {MonthlyReport.LeavesRowIndex.Row + 1} in sheet: {dataTable.TableName}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    hasReadErrors = true;
                                    logger.LogError($"Error on generating leave report for {monthlyReport}: {ex}");
                                    throw;
                                }
                            }
                            else
                            {
                                leaves[sheetName] = new int?();
                                totalLeaves = new int?();
                            }
                        });
                        logger.LogLine();
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
                string reportName = $"LeaveReport-FY{financialYear}";
                string sheetName = reportName;
                string fileName = $"{reportName}.xls";
                string exportPath = $"{exportFolder}\\{fileName}";

                DataTable dataTable = new(sheetName);
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
                    {
                        row[leaf.Key] = leaf.Value.HasValue ? leaf.Value.ToString() : "NA";
                    }

                    row["Total Leave Days"] = leaveReportData.TotalLeaves.HasValue ? leaveReportData.TotalLeaves.ToString() : "NA";
                    dataTable.Rows.Add(row);
                }
                WriteExcel(ref dataTable, exportPath, fileName);
            }
            else
            {
                logger.LogError("Process stopped due to errors.");
                return false;
            }
            return res;
        }

        public void ExportMonthlyReportInter(ref MonthlyReportData monthlyReportData, in string time, in string exportFolder)
        {
            logger.LogInfo("Exporting Monthly Report Inter:", 2);
            string reportName = "MonthlyReport_Inter";
            string sheetName = $"{reportName}_Sheet1";
            string fileName = $"{reportName}_{time}.xls";
            string exportPath = $"{exportFolder}\\{fileName}";
            DataTable monthlyTable = new(sheetName);
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
            foreach (IGrouping<string, EmployeeData>? grouping in monthlyReportData.EmployeesData.Where(employeeData => employeeData.ProjectData.Count > 0).GroupBy(ed => ed.Name).OrderBy(gp => gp.Key))
            {
                grouping.ToList().ForEach(employeeData =>
                {
                    employeeData.ProjectData.ToList().ForEach(pt =>
                    {
                        DataRow monthlyDtRow = monthlyTable.NewRow();
                        monthlyDtRow["Name(Id)"] = $"{employeeData.Name}({employeeData.Id})";
                        monthlyDtRow["Project Id"] = pt.Key;
                        monthlyDtRow["Actual Effort"] = DataService.TohhmmFormatString(pt.Value);
                        monthlyTable.Rows.Add(monthlyDtRow);
                    });
                });
            }
            WriteExcel(ref monthlyTable, exportPath, fileName);
        }

        public void ExportPtrInter(ref PtrData ptrData, in string time, in string exportFolder)
        {
            logger.LogInfo("Exporting PTR Inter:", 1);
            string reportName = "PTR_Inter";
            string sheetName = $"{reportName}_Sheet1";
            string fileName = $"{reportName}_{time}.xls";
            string exportPath = $"{exportFolder}\\{fileName}";
            DataTable dataTable = new(sheetName);
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
                row["Total Effort"] = DataService.TohhmmFormatString(projectEffort.Value);
                dataTable.Rows.Add(row);
            }
            WriteExcel(ref dataTable, exportPath, fileName);
        }

        public bool ExportPunchMovementSummaryReport(in string exportFolder, in string time, in PunchMovementData employeePunchData)
        {
            bool res = true;
            logger.LogInfo("Exporting Punch movement summary report:", 1);
            try
            {
                string reportName = "PunchMovementSummaryReport";
                string sheetName = $"{reportName}_Sheet1";
                string fileName = $"{reportName}_{time}.xls";
                string exportPath = $"{exportFolder}\\{fileName}";

                DataTable dataTable = new(sheetName);
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
                DataColumn column3 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Date",
                    Caption = "Date",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column3);
                DataColumn column4 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "First In",
                    Caption = "First In",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column4);
                DataColumn column5 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Last Out",
                    Caption = "Last Out",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column5);
                DataColumn column6 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "IsLastOutNextDay",
                    Caption = "IsLastOutNextDay",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column6);
                DataColumn column7 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Available Hours",
                    Caption = "Available Hours",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column7);
                DataColumn column8 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Work Hours",
                    Caption = "Work Hours",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column8);
                DataColumn column9 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Break Hours",
                    Caption = "Break Hours",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column9);

                void AddRow(PunchData punchData, DataRow row)
                {
                    row["Date"] = punchData.Date.ToShortDateString();
                    row["First In"] = punchData.FirstInTime.ToShortTimeString();
                    row["Last Out"] = punchData.LastOutTime.ToShortTimeString();
                    row["IsLastOutNextDay"] = punchData.IsLastOutNextDay.ToString();
                    row["Available Hours"] = DataService.TohhmmFormatString(punchData.AvailableHours);
                    row["Work Hours"] = DataService.TohhmmFormatString(punchData.WorkHours);
                    row["Break Hours"] = DataService.TohhmmFormatString(punchData.BreakHours);
                }

                for (int i = 0; i < employeePunchData.Length; i++)
                {
                    DataRow row = dataTable.NewRow();
                    row["Employee Id"] = employeePunchData[i].Id;
                    row["Employee Name"] = employeePunchData[i].Name;
                    AddRow(employeePunchData[i].PunchDatas[0], row);
                    dataTable.Rows.Add(row);

                    foreach (var punchData in employeePunchData[i].PunchDatas.Skip(1))
                    {
                        DataRow punchDataRow = dataTable.NewRow();
                        AddRow(punchData, punchDataRow);
                        dataTable.Rows.Add(punchDataRow);
                    }

                    DataRow summaryRow = dataTable.NewRow();
                    summaryRow["Date"] = "Summary";
                    summaryRow["Available Hours"] = DataService.TohhmmFormatString(employeePunchData[i].TotalAvailableHours);
                    summaryRow["Work Hours"] = DataService.TohhmmFormatString(employeePunchData[i].TotalWorkHours);
                    summaryRow["Break Hours"] = DataService.TohhmmFormatString(employeePunchData[i].TotalBreakHours);
                    dataTable.Rows.Add(summaryRow);
                }

                WriteExcel(ref dataTable, exportPath, fileName);
            }
            catch (Exception ex)
            {
                logger.LogError($"Error on exporting Punch movement summary report: {ex}");
                return false;
            }
            return res;
        }

        private void WriteExcel(ref DataTable dataTable, in string filePath, in string reportName = "data")
        {
            if (dataTable is not null && !string.IsNullOrEmpty(filePath))
            {
                if (!Directory.Exists(Path.GetDirectoryName(filePath)))
                {
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
                }

                using (StreamWriter streamWriter = new(filePath))
                {
                    for (int index = 0; index < dataTable.Columns.Count; ++index)
                    {
                        DefaultInterpolatedStringHandler interpolatedStringHandler = new(1, 1);
                        interpolatedStringHandler.AppendFormatted(dataTable.Columns[index]);
                        interpolatedStringHandler.AppendLiteral("\t");
                        string stringAndClear = interpolatedStringHandler.ToStringAndClear();
                        streamWriter.Write(stringAndClear);
                    }
                    streamWriter.WriteLine();
                    for (int index = 0; index < dataTable.Rows.Count; ++index)
                    {
                        for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; ++columnIndex)
                        {
                            streamWriter.Write(Convert.ToString(dataTable.Rows[index][columnIndex], CultureInfo.InvariantCulture) + "\t");
                        }

                        streamWriter.WriteLine();
                    }
                }
                logger.LogSameLine($"Exported {reportName}: ", 0); logger.LogDataSameLine(filePath);
            }
        }
    }
}