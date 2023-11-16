using ExcelDataReader;
using Models;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
using Utilities;

namespace Services
{
    public class ExportService(ILogger logger, DataService dataService)
    {
        public void ExportConsolidateData(in List<ConsolidatedData> consolidatedDataList, ref MonthlyReportData monthlyReportData, in string time, in string exportFolder)
        {
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
                    consDtRow["Total Effort as per PTR"] = dataService.TohhmmFormatString(consolidatedData.TotalEffort);
                    TotalEffort += consolidatedData.TotalEffort;
                    TimeSpan totalActualEffort = TimeSpan.Zero;
                    EmployeeActualEfforts.Where(eae => eae.ProjectId.Equals(consolidatedData.ProjectId, StringComparison.Ordinal)).ToList().ForEach(eae =>
                    {
                        totalActualEffort += eae.ActualEffort;
                        consDtRow[$"{eae.Name}({eae.Id})"] = dataService.TohhmmFormatString(eae.ActualEffort);
                    });
                    consDtRow["Total Actual Effort as per Monthly report"] = dataService.TohhmmFormatString(totalActualEffort);
                    TotalActualEffort += totalActualEffort;
                    dataTable.Rows.Add(consDtRow);
                }
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Total Hours";
                consDtRow["Total Effort as per PTR"] = dataService.TohhmmFormatString(TotalEffort);
                consDtRow["Total Actual Effort as per Monthly report"] = dataService.TohhmmFormatString(TotalActualEffort);
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                    .ForEach(ed =>
                    {
                        consDtRow[$"{ed.Name}({ed.Id})"] = dataService.TohhmmFormatString(ed.TotalProjectHours);
                    });
                dataTable.Rows.Add(consDtRow);
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "Actual Available Hours";
                consDtRow["Total Effort as per PTR"] = string.Empty;
                TimeSpan totalActualAvailableHours = TimeSpan.Zero;
                monthlyReportData.EmployeesData.Where(ed => employeeNames.Contains($"{ed.Name}({ed.Id})")).ToList()
                   .ForEach(ed =>
                   {
                       consDtRow[$"{ed.Name}({ed.Id})"] = dataService.TohhmmFormatString(ed.ActualAvailableHours);
                       totalActualAvailableHours += ed.ActualAvailableHours;
                   });
                consDtRow["Total Actual Effort as per Monthly report"] = dataService.TohhmmFormatString(totalActualAvailableHours);
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
                logger.LogErrorAndExit($"Error on exporting Consolidated Report: {ex}");
            }
        }

        public void ExportLeaveReport(in List<string> monthlyReports, in string financialYear, in string exportFolder)
        {
            logger.LogInfo("Generating Leave Report for FY" + financialYear + ":", 1);
            List<string> sheetNames = dataService.GetFyMonths(financialYear);
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
                    int? totalLeaves = new int?(0);
                    Dictionary<string, int?> leaves = [];
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
                                    logger.LogErrorAndExit($"Error on generating leave report for {monthlyReport}: {ex}");
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
                logger.LogErrorAndExit("Process stopped due to errors.");
            }
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
                        monthlyDtRow["Actual Effort"] = dataService.TohhmmFormatString(pt.Value);
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
                row["Total Effort"] = dataService.TohhmmFormatString(projectEffort.Value);
                dataTable.Rows.Add(row);
            }
            WriteExcel(ref dataTable, exportPath, fileName);
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