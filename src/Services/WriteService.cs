﻿using Models;
using Spire.Xls;
using System.Globalization;
using Utilities;

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

                    var sheetNames = workbook.Worksheets.Select(x => x.Name);

                    List<(string sheetName, int year, int month)> dataEntrySheets = [];

                    foreach (var sheetName in sheetNames)
                    {
                        var splits = sheetName.Trim().Split('_');

                        if (splits.Length != 2) continue;

                        var monthString = splits[0];
                        var yearString = splits[1];

                        bool isMonthParsable = DateTime.TryParseExact(monthString, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime monthDateTime);
                        if (!isMonthParsable) continue;
                        var month = monthDateTime.Month;

                        bool isYearParsable = DateTime.TryParseExact(yearString, "YYYY", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime yearDateTime);
                        if (!isYearParsable) continue;
                        var year = yearDateTime.Year;

                        dataEntrySheets.Add((sheetName, year, month));
                    }

                    logger.LogSameLine($"Editing Sheet");
                    foreach ((string sheetName, int year, int month) in dataEntrySheets)
                    {
                        logger.LogDataSameLine(Path.GetFileName(sheetName));

                        var workSheet = workbook.Worksheets[sheetName];
                        //TODO: 1. Get unique employee ids from sheet
                        //TODO: 2. get matching muster data
                        //TODO: 3. get entry row indexes
                        //TODO: 4. Enter muster values
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
                        logger.LogWarning($"Ignoring monthly report {fileName}, unable to locate 5 digit employee id in the file name");
                        continue;
                    }

                    Workbook workbook = new();

                    workbook.LoadFromFile(fileName);

                    musterOptionsDatas.Datas.TryGetValue(fileId, out var musterOptionsData);

                    if (musterOptionsData == null) continue;

                    var dataDates = musterOptionsData.MusterOptions.Select(x => x.Date).DistinctBy(x => x.Date.Month).ToList();

                    foreach (var dataDate in dataDates)
                    {
                        var sheetName = dataDate.ToString("MMM-yy");

                        Worksheet worksheet = workbook.Worksheets[sheetName];

                        var musterOptions = musterOptionsData.MusterOptions
                            .Where(x => x.Date.Month == dataDate.Month)
                            .OrderBy(x => x.Date)
                            .ToArray();

                        int inTimeRowIndex = 10;
                        int outTimeRowIndex = 12;
                        int firstDateColumnIndex = 5;
                        int lastDateColumnIndex = firstDateColumnIndex + musterOptions.Length;
                        int dataIndex = 0;
                        for (int i = firstDateColumnIndex; i < lastDateColumnIndex; i++)
                        {
                            var musterOption = musterOptions[dataIndex];

                            var inTime = musterOption?.InTime;
                            if (inTime != null)
                            {
                                worksheet.SetCellValue(inTimeRowIndex, i, $"{inTime.Value.Hour}:{inTime.Value.Minute}");
                                worksheet.CellList[inTimeRowIndex].Style.NumberFormat = "[h]:mm";
                            }

                            var outTime = musterOption?.OutTime;
                            if (outTime != null)
                            {
                                worksheet.SetCellValue(outTimeRowIndex, i, $"{outTime.Value.Hour}:{outTime.Value.Minute}");
                                worksheet.CellList[outTimeRowIndex].Style.NumberFormat = "[h]:mm";
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