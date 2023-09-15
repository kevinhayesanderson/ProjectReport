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
            ConsoleLogger.LogInfo("Exporting consolidated data.", 1);
            try
            {
                string tableName = "ConsolidatedReport_" + time;
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
                    DataType = typeof(double),
                    ColumnName = "Total Effort",
                    Caption = "Total Effort",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column2);
                DataColumn column3 = new()
                {
                    DataType = typeof(string),
                    ColumnName = "Total Actual Effort",
                    Caption = "Total Actual Effort",
                    ReadOnly = false,
                    Unique = false
                };
                dataTable.Columns.Add(column3);
                List<EmployeeActualEffort> list = consolidatedDataList.SelectMany(data => (IEnumerable<EmployeeActualEffort>)data.EmployeeActualEffort).ToList();
                List<string> employeeNames = list.Select(eae =>
                {
                    DefaultInterpolatedStringHandler interpolatedStringHandler = new(2, 2);
                    interpolatedStringHandler.AppendFormatted(eae.Name);
                    interpolatedStringHandler.AppendLiteral("(");
                    interpolatedStringHandler.AppendFormatted(eae.Id);
                    interpolatedStringHandler.AppendLiteral(")");
                    return interpolatedStringHandler.ToStringAndClear();
                }).Distinct().ToList();
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
                double num = 0.0;
                TimeSpan timeSpan1 = new();
                DataRow consDtRow;
                foreach (ConsolidatedData consolidatedData1 in consolidatedDataList)
                {
                    ConsolidatedData consolidatedData = consolidatedData1;
                    consDtRow = dataTable.NewRow();
                    consDtRow["Project Id"] = consolidatedData.ProjectId;
                    consDtRow["Total Effort"] = consolidatedData.TotalEffort;
                    num += consolidatedData.TotalEffort;
                    TimeSpan totalActualEffort = new();
                    list.Where(eae => eae.ProjectId.Equals(consolidatedData.ProjectId, StringComparison.Ordinal)).ToList().ForEach(eae =>
                    {
                        totalActualEffort += eae.ActualEffort;
                        DataRow dataRow = consDtRow;
                        DefaultInterpolatedStringHandler interpolatedStringHandler = new(2, 2);
                        interpolatedStringHandler.AppendFormatted(eae.Name);
                        interpolatedStringHandler.AppendLiteral("(");
                        interpolatedStringHandler.AppendFormatted(eae.Id);
                        interpolatedStringHandler.AppendLiteral(")");
                        string stringAndClear1 = interpolatedStringHandler.ToStringAndClear();
                        interpolatedStringHandler = new DefaultInterpolatedStringHandler(1, 2);
                        interpolatedStringHandler.AppendFormatted((int)eae.ActualEffort.TotalHours);
                        interpolatedStringHandler.AppendLiteral(":");
                        interpolatedStringHandler.AppendFormatted(eae.ActualEffort.Minutes);
                        string stringAndClear2 = interpolatedStringHandler.ToStringAndClear();
                        dataRow[stringAndClear1] = stringAndClear2;
                    });
                    DataRow dataRow1 = consDtRow;
                    DefaultInterpolatedStringHandler interpolatedStringHandler1 = new(1, 2);
                    interpolatedStringHandler1.AppendFormatted((int)totalActualEffort.TotalHours);
                    interpolatedStringHandler1.AppendLiteral(":");
                    interpolatedStringHandler1.AppendFormatted(totalActualEffort.Minutes);
                    string stringAndClear = interpolatedStringHandler1.ToStringAndClear();
                    dataRow1["Total Actual Effort"] = stringAndClear;
                    timeSpan1 += totalActualEffort;
                    dataTable.Rows.Add(consDtRow);
                }
                consDtRow = dataTable.NewRow();
                consDtRow["Project Id"] = "TOTAL HOURS";
                consDtRow["Total Effort"] = num;
                DataRow dataRow2 = consDtRow;
                DefaultInterpolatedStringHandler interpolatedStringHandler2 = new(1, 2);
                interpolatedStringHandler2.AppendFormatted((int)timeSpan1.TotalHours);
                interpolatedStringHandler2.AppendLiteral(":");
                interpolatedStringHandler2.AppendFormatted(timeSpan1.Minutes);
                string stringAndClear3 = interpolatedStringHandler2.ToStringAndClear();
                dataRow2["Total Actual Effort"] = stringAndClear3;
                monthlyReportData.EmployeesData.Where(ed =>
                {
                    List<string> stringList = employeeNames;
                    DefaultInterpolatedStringHandler interpolatedStringHandler3 = new(2, 2);
                    interpolatedStringHandler3.AppendFormatted(ed.Name);
                    interpolatedStringHandler3.AppendLiteral("(");
                    interpolatedStringHandler3.AppendFormatted(ed.Id);
                    interpolatedStringHandler3.AppendLiteral(")");
                    string stringAndClear4 = interpolatedStringHandler3.ToStringAndClear();
                    return stringList.Contains(stringAndClear4);
                }).ToList().ForEach(ed =>
                {
                    TimeSpan timeSpan2 = ed.ProjectTime.Join((IEnumerable<KeyValuePair<string, double>>)ptrData.ProjectEfforts, projTime => projTime.Key, projEffort => projEffort.Key, (projTime, projEffort) => projTime.Value).ToList().Aggregate(new TimeSpan(), (current, item) => current + item);
                    DataRow dataRow3 = consDtRow;
                    DefaultInterpolatedStringHandler interpolatedStringHandler4 = new(2, 2);
                    interpolatedStringHandler4.AppendFormatted(ed.Name);
                    interpolatedStringHandler4.AppendLiteral("(");
                    interpolatedStringHandler4.AppendFormatted(ed.Id);
                    interpolatedStringHandler4.AppendLiteral(")");
                    string stringAndClear5 = interpolatedStringHandler4.ToStringAndClear();
                    interpolatedStringHandler4 = new DefaultInterpolatedStringHandler(1, 2);
                    interpolatedStringHandler4.AppendFormatted((int)timeSpan2.TotalHours);
                    interpolatedStringHandler4.AppendLiteral(":");
                    interpolatedStringHandler4.AppendFormatted(timeSpan2.Minutes);
                    string stringAndClear6 = interpolatedStringHandler4.ToStringAndClear();
                    dataRow3[stringAndClear5] = stringAndClear6;
                });
                dataTable.Rows.Add(consDtRow);
                WriteExcel(dataTable, exportPath);
            }
            catch (Exception ex)
            {
                ConsoleLogger.LogErrorAndExit("Error on exporting data: " + ex.Message);
            }
        }

        public static void ExportLeaveReport(List<string> monthlyReports, string financialYear, string time, string exportFolder)
        {
            ConsoleLogger.LogInfo("Generating Leave Report for FY" + financialYear + ".", 1);
            List<string> sheetNames = DataService.GetFyMonths(financialYear);
            List<LeaveReportData> leaveReportDataList = new();
            bool hasReadErrors = false;
            foreach (string monthlyReport1 in monthlyReports)
            {
                string monthlyReport = monthlyReport1;
                LeaveReportData leaveReportData;
                using (FileStream fileStream = File.Open(monthlyReport, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream, null))
                    {
                        DataTableCollection tables = ExcelDataReaderExtensions.AsDataSet(reader, null).Tables;
                        List<DataTable> sheets = tables.Cast<DataTable>().Where(dataTable => sheetNames.Contains(dataTable.TableName.Trim())).ToList();
                        int? totalLeaves = new int?(0);
                        Dictionary<string, int?> leaves = new();
                        string str;
                        int result1;
                        if (sheets.Count > 0)
                        {
                            str = (string)sheets[0].Rows[3][2];
                            if (string.IsNullOrEmpty(str))
                            {
                                DefaultInterpolatedStringHandler interpolatedStringHandler = new(69, 3);
                                interpolatedStringHandler.AppendLiteral("Name cannot be null or empty string in sheet ");
                                interpolatedStringHandler.AppendFormatted(sheets[0].TableName);
                                interpolatedStringHandler.AppendLiteral(": Check row ");
                                interpolatedStringHandler.AppendFormatted(4);
                                interpolatedStringHandler.AppendLiteral(" at column ");
                                interpolatedStringHandler.AppendFormatted(3);
                                interpolatedStringHandler.AppendLiteral(".");
                                throw new ArgumentException(interpolatedStringHandler.ToStringAndClear());
                            }
                            if (!int.TryParse(sheets[0].Rows[4][2].ToString(), out result1))
                            {
                                DefaultInterpolatedStringHandler interpolatedStringHandler = new(45, 3);
                                interpolatedStringHandler.AppendLiteral("Invalid format at column ");
                                interpolatedStringHandler.AppendFormatted(3);
                                interpolatedStringHandler.AppendLiteral(": Check row ");
                                interpolatedStringHandler.AppendFormatted(5);
                                interpolatedStringHandler.AppendLiteral(" sheet ");
                                interpolatedStringHandler.AppendFormatted(sheets[0].TableName);
                                interpolatedStringHandler.AppendLiteral(".");
                                throw new FormatException(interpolatedStringHandler.ToStringAndClear());
                            }
                            sheetNames.ForEach(sheetName =>
                            {
                                if (sheets.Any(sh => sh.TableName == sheetName))
                                {
                                    DataTable dataTable = sheets.First(sh => sh.TableName == sheetName);
                                    try
                                    {
                                        DataRowCollection rows = dataTable.Rows;
                                        int columnIndex = rows[14].ItemArray.Length - 1;
                                        int result2;
                                        if (int.TryParse(rows[14][columnIndex].ToString(), out result2))
                                        {
                                            leaves[dataTable.TableName] = new int?(result2);
                                            int? nullable = totalLeaves;
                                            int num = result2;
                                            totalLeaves = nullable.HasValue ? new int?(nullable.GetValueOrDefault() + num) : new int?();
                                        }
                                        else
                                        {
                                            DefaultInterpolatedStringHandler interpolatedStringHandler = new(50, 2);
                                            interpolatedStringHandler.AppendLiteral("Invalid format at column ");
                                            interpolatedStringHandler.AppendFormatted(columnIndex + 1);
                                            interpolatedStringHandler.AppendLiteral(": Check row 15 in sheet ");
                                            interpolatedStringHandler.AppendFormatted(dataTable.TableName);
                                            interpolatedStringHandler.AppendLiteral(".");
                                            throw new FormatException(interpolatedStringHandler.ToStringAndClear());
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
                        }
                        else
                        {
                            str = string.IsNullOrEmpty(tables[0].Rows[3][2].ToString()) ? "NAME NOT PROVIDED AT SHEET 1" : (string)tables[0].Rows[3][2];
                            if (!int.TryParse(tables[0].Rows[4][2].ToString(), out result1))
                                result1 = 0;
                            sheetNames.ForEach(sheetName => leaves[sheetName] = new int?());
                            totalLeaves = new int?();
                        }
                        leaveReportData = new LeaveReportData()
                        {
                            EmpCode = result1.ToString(CultureInfo.InvariantCulture),
                            Name = str,
                            Leaves = leaves,
                            TotalLeaves = totalLeaves
                        };
                    }
                }
                leaveReportDataList.Add(leaveReportData);
            }
            if (!hasReadErrors)
            {
                string tableName = "LeaveReport-FY" + financialYear + "_" + time;
                string exportPath = exportFolder + "\\" + tableName + ".xls";
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
                    row["Employee Id"] = leaveReportData.EmpCode;
                    row["Employee Name"] = leaveReportData.Name;
                    foreach (KeyValuePair<string, int?> leaf in leaveReportData.Leaves)
                        row[leaf.Key] = !leaf.Value.HasValue ? "NA" : (object)leaf.Value.ToString();
                    row["Total Leave Days"] = !leaveReportData.TotalLeaves.HasValue ? "NA" : (object)leaveReportData.TotalLeaves.ToString();
                    dataTable.Rows.Add(row);
                }
                WriteExcel(dataTable, exportPath);
            }
            else
                ConsoleLogger.LogErrorAndExit("Process stopped due to errors.");
        }

        public static void ExportMonthlyReportInter(MonthlyReportData monthlyReportData, string time, string exportFolder)
        {
            ConsoleLogger.LogInfo("Exporting monthly report inter data.", 2);
            string tableName = "MonthlyReport_Inter_" + time;
            string exportPath = exportFolder + "\\" + tableName + ".xls";
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
            foreach (EmployeeData employeeData1 in monthlyReportData.EmployeesData.Where<EmployeeData>(employeeData => employeeData.ProjectTime.Count > 0))
            {
                EmployeeData employeeData = employeeData1;
                employeeData.ProjectTime.ToList().ForEach(pt =>
                {
                    DataRow monthlyDtRow = monthlyTable.NewRow();
                    DataRow dataRow1 = monthlyDtRow;
                    DefaultInterpolatedStringHandler interpolatedStringHandler = new(2, 2);
                    interpolatedStringHandler.AppendFormatted(employeeData.Name);
                    interpolatedStringHandler.AppendLiteral("(");
                    interpolatedStringHandler.AppendFormatted(employeeData.Id);
                    interpolatedStringHandler.AppendLiteral(")");
                    string stringAndClear1 = interpolatedStringHandler.ToStringAndClear();
                    dataRow1["Name(Id)"] = stringAndClear1;
                    monthlyDtRow["Project Id"] = pt.Key;
                    DataRow dataRow2 = monthlyDtRow;
                    interpolatedStringHandler = new DefaultInterpolatedStringHandler(1, 2);
                    interpolatedStringHandler.AppendFormatted((int)pt.Value.TotalHours);
                    interpolatedStringHandler.AppendLiteral(":");
                    interpolatedStringHandler.AppendFormatted(pt.Value.Minutes);
                    string stringAndClear2 = interpolatedStringHandler.ToStringAndClear();
                    dataRow2["Actual Effort"] = stringAndClear2;
                    monthlyTable.Rows.Add(monthlyDtRow);
                });
            }
            WriteExcel(monthlyTable, exportPath);
        }

        public static void ExportPtrInter(PtrData ptrData, string time, string exportFolder)
        {
            ConsoleLogger.LogInfo("Exporting ptr inter data.", 1);
            string tableName = "PTR_Inter_" + time;
            string exportPath = exportFolder + "\\" + tableName + ".xls";
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
            foreach (KeyValuePair<string, double> projectEffort in ptrData.ProjectEfforts)
            {
                DataRow row = dataTable.NewRow();
                row["Project Id"] = projectEffort.Key;
                row["Total Effort"] = projectEffort.Value;
                dataTable.Rows.Add(row);
            }
            WriteExcel(dataTable, exportPath);
        }

        private static void WriteExcel(DataTable dataTable, string exportPath)
        {
            DirectoryInfo directoryInfo = new(exportPath);
            if (!Directory.Exists(Path.GetDirectoryName(exportPath)))
                directoryInfo = Directory.CreateDirectory(Path.GetDirectoryName(exportPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            using (StreamWriter streamWriter1 = new(exportPath))
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
            ConsoleLogger.Log("Exported data to " + directoryInfo.FullName + ".", 1);
        }
    }
}