using Models;
using Spire.Xls;
using System.Globalization;
using Utilities;
using static Models.Constants;

namespace Services
{
    public class WriteService(ILogger logger)
    {
        public bool WriteAttendanceReportEntry(in List<string> attendanceReports, in MusterOptionsDatas musterOptionsDatas)
        {
            bool res = true;
            logger.LogInfo("Writing AttendanceReportEntry:", 2);

            try
            {
                foreach (var attendanceReport in attendanceReports)
                {
                    logger.LogSameLine("Editing Attendance Report");
                    logger.LogDataSameLine(Path.GetFileName(attendanceReport));
                    logger.LogLine();

                    Workbook workbook = new();

                    workbook.LoadFromFile(attendanceReport);

                    var sheetNames = workbook.Worksheets.Select(x => x.Name.Trim());

                    List<(string sheetName, DateOnly dateOnly)> dataEntrySheets = [];

                    foreach (var sheetName in sheetNames)
                    {
                        if (DateOnly.TryParseExact(sheetName, AttendanceReport.SheetNamePattern, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly dateOnly))
                        {
                            dataEntrySheets.Add((sheetName, dateOnly));
                        }
                    }

                    logger.LogSameLine($"Editing Sheet");
                    foreach ((string sheetName, DateOnly dateOnly) in dataEntrySheets)
                    {
                        logger.LogDataSameLine($"{sheetName} ,");

                        var workSheet = workbook.Worksheets[sheetName];

                        List<CellRange> empCodeColumnCellList = workSheet.Columns[AttendanceReport.EmpCodeIndex.Column].CellList;
                        var uniqueEmployeeIds = empCodeColumnCellList
                            .Where(x => x.NumberValue != double.NaN && uint.TryParse(x.NumberText, out _))
                            .Select(x => uint.Parse(x.NumberText))
                            .ToHashSet();

                        var musterDatas = musterOptionsDatas.Datas
                            .Where(x => uniqueEmployeeIds.Contains(x.Key) && x.Value.MusterOptions.Exists(y => y.Date.Month == dateOnly.Month && y.Date.Year == dateOnly.Year))
                            .ToDictionary();

                        var entryRowData = uniqueEmployeeIds
                            .Select(uniqueEmployeeId => (uniqueEmployeeId, empCodeColumnCellList.FindIndex(x => uint.TryParse(x.NumberText, out uint numberValue) && numberValue == uniqueEmployeeId)))
                            .ToList();

                        foreach ((uint uniqueEmployeeId, int rowIndex) in entryRowData)
                        {
                            int column = AttendanceReport.DateStartIndex.Column;
                            if (musterDatas.TryGetValue(uniqueEmployeeId, out MusterOptionsData musterOptionsData))
                            {
                                foreach (var musterOption in musterOptionsData.MusterOptions.Where(x => x.Date.Month == dateOnly.Month && x.Date.Year == dateOnly.Year))
                                {
                                    if (!string.IsNullOrEmpty(musterOption.Shift.Trim())) workSheet.SetCellValue(rowIndex + 2, column, musterOption.Shift);
                                    if (musterOption.InTime is TimeOnly inTime)
                                    {
                                        workSheet.SetCellValue(rowIndex + 3, column, $"{inTime.Hour}:{inTime.Minute}");
                                        workSheet[rowIndex + 3, column].Style.NumberFormat = AttendanceReport.TimeNumberFormat;
                                    }
                                    if (musterOption.OutTime is TimeOnly outTime)
                                    {
                                        workSheet.SetCellValue(rowIndex + 4, column, $"{outTime.Hour}:{outTime.Minute}");
                                        workSheet[rowIndex + 4, column].Style.NumberFormat = AttendanceReport.TimeNumberFormat;
                                    }
                                    if (!string.IsNullOrEmpty(musterOption.Muster.Trim())) workSheet.SetCellValue(rowIndex + 5, column, musterOption.Muster);
                                    column++;
                                }
                            }
                        }
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"An error occurred on writing AttendanceReportEntry: {ex.Message}");
                return false;
            }
            return res;
        }

        public bool WriteMonthlyReportInOutEntry(in List<(uint, string)> monthlyReportsData, in MusterOptionsDatas musterOptionsDatas)
        {
            bool res = true;
            logger.LogInfo("Writing MonthlyReportInOutEntry:", 2);

            try
            {
                foreach ((uint fileId, string fileName) in monthlyReportsData)
                {
                    logger.LogSameLine("Editing ");
                    logger.LogDataSameLine(fileName);
                    logger.LogLine();

                    if (fileId == uint.MinValue)
                    {
                        logger.LogWarning($"Ignoring monthly report {fileName}, unable to locate {General.EmployeeIdLength} digit employee id in the file name");
                        continue;
                    }

                    Workbook workbook = new();

                    workbook.LoadFromFile(fileName);

                    musterOptionsDatas.Datas.TryGetValue(fileId, out var musterOptionsData);

                    if (musterOptionsData == null) continue;

                    var dataDates = musterOptionsData.MusterOptions.Select(x => x.Date).DistinctBy(x => x.Date.Month).ToList();

                    foreach (var dataDate in dataDates)
                    {
                        var sheetName = dataDate.ToString(MonthlyReport.SheetNamePattern);

                        Worksheet worksheet = workbook.Worksheets[sheetName];

                        var musterOptions = musterOptionsData.MusterOptions
                            .Where(x => x.Date.Month == dataDate.Month)
                            .OrderBy(x => x.Date)
                            .ToArray();

                        int lastDateColumnIndex = MonthlyReport.FirstDateIndex.Column + musterOptions.Length;
                        int dataIndex = 0;
                        for (int i = MonthlyReport.FirstDateIndex.Column; i < lastDateColumnIndex; i++)
                        {
                            var musterOption = musterOptions[dataIndex];

                            var inTime = musterOption?.InTime;
                            if (inTime != null)
                            {
                                worksheet.SetCellValue(MonthlyReport.InTimeIndex.Row, i, $"{inTime.Value.Hour}:{inTime.Value.Minute}");
                                worksheet.CellList[MonthlyReport.InTimeIndex.Row].Style.NumberFormat = MonthlyReport.TimeNumberFormat;
                            }

                            var outTime = musterOption?.OutTime;
                            if (outTime != null)
                            {
                                worksheet.SetCellValue(MonthlyReport.OutTimeIndex.Row, i, $"{outTime.Value.Hour}:{outTime.Value.Minute}");
                                worksheet.CellList[MonthlyReport.OutTimeIndex.Row].Style.NumberFormat = MonthlyReport.TimeNumberFormat;
                            }

                            dataIndex++;
                        }
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"An error occurred on writing MonthlyReportInOutEntry: {ex.Message}");
                return false;
            }
            return res;
        }
    }
}