// Ignore Spelling: Tohhmm

using Models;
using Utilities;

namespace Services
{
    public class DataService(ILogger logger)
    {
        public long[] Months => [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

        public void CalculatePunchMovement(PunchMovementData punchMovementData, string cutOff)
        {
            foreach (var punchDatas in punchMovementData.EmployeePunchDatas.Select(epd => epd.PunchDatas))
            {
                for (int i = 0; i < punchDatas.Count; i++)
                {
                    (int hour, int minute) = (int.Parse(cutOff.Split(':')[1]), int.Parse(cutOff.Split(':')[^1]));

                    DateTime firstValue = punchDatas[i].Punches.First(punch => TimeOnly.FromDateTime(punch) >= new TimeOnly(hour, minute));

                    IEnumerable<DateTime> InOuts = Enumerable.Empty<DateTime>();

                    bool lastOutPredicate(DateTime punch) => TimeOnly.FromDateTime(punch) <= new TimeOnly(hour, minute);

                    bool TryGetLastOutForPreviousDay(List<DateTime> punches, out DateTime lastOut)
                    {
                        bool res = false;
                        lastOut = DateTime.MinValue;
                        res = punches.Exists(lastOutPredicate);
                        if (res) lastOut = punches.Last(lastOutPredicate);
                        return res;
                    }

                    bool isLastOutNextDay = false;

                    DateTime lastValue;
                    if ((i + 1) < punchDatas.Count && TryGetLastOutForPreviousDay(punchDatas[i + 1].Punches, out DateTime lastOut))
                    {
                        isLastOutNextDay = true;
                        lastValue = lastOut;
                        int subsetIndex = punchDatas[i + 1].Punches.IndexOf(lastValue);
                        var nextDayInOuts = punchDatas[i + 1].Punches[..subsetIndex];
                        var firstValueIndex = punchDatas[i].Punches.IndexOf(firstValue);
                        InOuts = punchDatas[i].Punches[firstValueIndex..].Concat(nextDayInOuts);
                    }
                    else
                    {
                        lastValue = punchDatas[i].Punches[^1];
                        var firstIndex = punchDatas[i].Punches.IndexOf(firstValue);
                        var lastIndex = punchDatas[i].Punches.IndexOf(lastValue);
                        InOuts = punchDatas[i].Punches[firstIndex..lastIndex];
                    }

                    TimeSpan breakTime = TimeSpan.MinValue;

                    if (InOuts != null && (InOuts.Count() % 2 == 0))
                    {
                        var ins = InOuts.Where(io => Array.IndexOf(InOuts.ToArray(), io) % 2 == 0);
                        var outs = InOuts.Except(ins);
                        foreach ((DateTime outTime, DateTime inTime) in outs.Zip(ins))
                        {
                            breakTime.Add(TimeOnly.FromDateTime(outTime) - TimeOnly.FromDateTime(inTime));
                        }
                    }

                    punchDatas[i] = new PunchData
                    {
                        Date = punchDatas[i].Date,
                        Punches = punchDatas[i].Punches,
                        FirstInTime = firstValue,
                        LastOutTime = lastValue,
                        WorkHours = TimeOnly.FromDateTime(firstValue) - TimeOnly.FromDateTime(lastValue),
                        BreakHours = breakTime,
                        IsLastOutNextDay = isLastOutNextDay
                    };
                }
            }
        }

        public List<ConsolidatedData> Consolidate(PtrData ptrData, MonthlyReportData monthlyReportData)
        {
            List<ConsolidatedData> consolidatedDataList = [];
            try
            {
                logger.LogInfo("Consolidating data", 2);
                IEnumerable<string> projectIds = ptrData.ProjectIds.Union(monthlyReportData.ProjectIds);
                consolidatedDataList.AddRange(from projectId in projectIds
                                              let consolidatedData = new ConsolidatedData
                                              {
                                                  ProjectId = projectId,
                                                  TotalEffort = ptrData.ProjectEfforts.TryGetValue(projectId, out TimeSpan value) ? value : TimeSpan.Zero,
                                                  EmployeeActualEffort = monthlyReportData.EmployeesData
                                                                        .Where(ed => ed.ProjectData.ContainsKey(projectId))
                                                                        .Select(ed => new EmployeeActualEffort()
                                                                        {
                                                                            Id = ed.Id,
                                                                            Name = ed.Name,
                                                                            ProjectId = projectId,
                                                                            ActualEffort = ed.ProjectData[projectId]
                                                                        }).ToList()
                                              }
                                              select consolidatedData);
            }
            catch (Exception ex)
            {
                logger.LogError($"Error on consolidating data: {ex} ");
                throw;
            }
            return consolidatedDataList;
        }

        public List<string> GetFyMonths(string financialYear)
        {
            string[] strArray = financialYear.Split('-');
            return
            [
                $"Apr-{strArray[0]}",
                $"May-{strArray[0]}",
                $"Jun-{strArray[0]}",
                $"Jul-{strArray[0]}",
                $"Aug-{strArray[0]}",
                $"Sep-{strArray[0]}",
                $"Oct-{strArray[0]}",
                $"Nov-{strArray[0]}",
                $"Dec-{strArray[0]}",
                $"Jan-{strArray[1]}",
                $"Feb-{strArray[1]}",
                $"Mar-{strArray[1]}"
            ];
        }

        public string TohhmmFormatString(TimeSpan timeSpan)
        {
            string totalHours = ((int)timeSpan.TotalHours).ToString();
            string minutes = timeSpan.Minutes.ToString();
            if (timeSpan.Seconds == 59 && (timeSpan.Milliseconds == 999 || timeSpan.Microseconds == 999))
            {
                if (timeSpan.Minutes + 1 == 60)
                {
                    minutes = "00";
                    totalHours = ((int)timeSpan.TotalHours + 1).ToString();
                }
                else
                {
                    minutes = (timeSpan.Minutes + 1).ToString();
                }
            }
            return $"{totalHours}:{minutes}";
        }
    }
}