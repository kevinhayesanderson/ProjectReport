using Models;
using Spire.Xls;
using Utilities;

namespace Services
{
    public class WriteService(ILogger logger)
    {
        public bool WriteInOutEntry(List<(uint, string)> monthlyReportsData, MusterOptionsDatas musterOptionsDatas)
        {
            bool res = true;
            logger.LogInfo("Writing InOutEntry in monthly reports:", 2);

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
                        int firtDateColumnIndex = 4;
                        int lastDateColumnIndex = firtDateColumnIndex + musterOptions.Length;
                        int dataIndex = 0;
                        for (int i = firtDateColumnIndex; i < lastDateColumnIndex; i++)
                        {
                            var musterOption = musterOptions[dataIndex];

                            var inTime = musterOption?.InTime;
                            if (inTime != null)
                                worksheet.SetCellValue(inTimeRowIndex, i, inTime.Value.ToString("H:mm"));

                            var outTime = musterOption?.OutTime;
                            if (outTime != null)
                                worksheet.SetCellValue(outTimeRowIndex, i, outTime.Value.ToString("H:mm"));

                            dataIndex++;
                        }
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"An error occurred on writing InOutEntry in monthly reports: {ex.Message}");
                return false;
            }
            return res;
        }
    }
}