// Ignore Spelling: Tohhmm

using Models;
using Utilities;

namespace Services
{
    public class DataService(ILogger logger)
    {
        public long[] Months => [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

        public static string TohhmmFormatString(TimeSpan timeSpan)
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

        public void CalculatePunchMovement(PunchMovementData punchMovementData, string cutOff)
        {
            logger.LogInfo("Calculating punch movement...", 2);
            try
            {
                foreach (var employeePunchDatas in punchMovementData.EmployeePunchDatas)
                {
                    var punchDatas = employeePunchDatas.PunchDatas;
                    for (int i = 0; i < punchDatas.Count; i++)
                    {
                        (int cutOffHour, int cutOffMinute) = (int.Parse(cutOff.Split(':')[0]), int.Parse(cutOff.Split(':')[^1]));

                        TimeOnly firstValue = punchDatas[i].Punches.Find(punch => punch >= new TimeOnly(cutOffHour, cutOffMinute));

                        int firstValueIndex = punchDatas[i].Punches.IndexOf(firstValue);

                        if (firstValue == default || (firstValueIndex != -1 && punchDatas[i].Punches[firstValueIndex..].Count <= 1))
                        {
                            continue;
                        }

                        IEnumerable<TimeOnly> InOuts = Enumerable.Empty<TimeOnly>();

                        bool lastOutPredicate(TimeOnly punch) => punch <= new TimeOnly(cutOffHour, cutOffMinute);

                        bool TryGetLastOutInNextDay(List<TimeOnly> punches, out TimeOnly lastOut)
                        {
                            bool res = false;
                            lastOut = TimeOnly.MinValue;
                            res = punches.Exists(lastOutPredicate);
                            if (res) lastOut = punches.Last(lastOutPredicate);
                            return res;
                        }

                        bool isLastOutNextDay = false;

                        TimeOnly lastValue;
                        if ((i + 1) < punchDatas.Count && TryGetLastOutInNextDay(punchDatas[i + 1].Punches, out TimeOnly lastOut))
                        {
                            isLastOutNextDay = true;
                            lastValue = lastOut;
                            int subsetIndex = punchDatas[i + 1].Punches.IndexOf(lastValue);
                            var nextDayInOuts = punchDatas[i + 1].Punches[..(subsetIndex + 1)];
                            InOuts = punchDatas[i].Punches[firstValueIndex..].Concat(nextDayInOuts);
                        }
                        else
                        {
                            lastValue = punchDatas[i].Punches[^1];
                            var lastValueIndex = punchDatas[i].Punches.IndexOf(lastValue);
                            InOuts = punchDatas[i].Punches[firstValueIndex..(lastValueIndex + 1)];
                        }

                        TimeSpan breakTime = TimeSpan.Zero;

                        if (InOuts != null)
                        {
                            if (InOuts.Count() % 2 != 0)
                            {
                                var lastIn = InOuts.Last();
                                var lastEvenOut = InOuts.ToArray()[InOuts.Count() - 2];
                                InOuts = InOuts.ToArray()[..^1].Concat([lastEvenOut, lastIn]);
                            }
                            var ins = InOuts.Where(io => Array.IndexOf(InOuts.ToArray(), io) % 2 == 0).ToArray();
                            var outs = InOuts.Except(ins).ToArray();
                            foreach ((TimeOnly outTime, TimeOnly inTime) in outs.Zip(ins[1..]))
                            {
                                breakTime += inTime - outTime;
                            }
                        }

                        punchDatas[i] = new PunchData
                        {
                            Date = punchDatas[i].Date,
                            Punches = punchDatas[i].Punches,
                            FirstInTime = firstValue,
                            LastOutTime = lastValue,
                            AvailableHours = lastValue - firstValue,
                            BreakHours = breakTime,
                            WorkHours = lastValue - firstValue - breakTime,
                            IsLastOutNextDay = isLastOutNextDay
                        };
                    }

                    var totalAvailableHours = TimeSpan.Zero;
                    var totalWorkHours = TimeSpan.Zero;
                    var totalBreakHours = TimeSpan.Zero;
                    foreach (var punchData in employeePunchDatas.PunchDatas)
                    {
                        totalAvailableHours += punchData.AvailableHours;
                        totalWorkHours += punchData.WorkHours;
                        totalBreakHours += punchData.BreakHours;
                    }
                    employeePunchDatas.TotalAvailableHours = totalAvailableHours;
                    employeePunchDatas.TotalWorkHours = totalWorkHours;
                    employeePunchDatas.TotalBreakHours = totalBreakHours;
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Error on calculating punch movement: {ex} ");
                throw;
            }
        }

        public List<ConsolidatedData> Consolidate(PtrData ptrData, MonthlyReportData monthlyReportData)
        {
            List<ConsolidatedData> consolidatedDataList = [];
            try
            {
                logger.LogInfo("Consolidating data...", 2);
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
    }
}