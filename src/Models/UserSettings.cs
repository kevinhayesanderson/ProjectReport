using System.Text.Json.Serialization;

namespace Models
{
    public class UserSettings
    {
        [JsonPropertyName("Actions")]
        public Action[] Actions { get; set; } = [];
    }

    public struct UserSettingsInfo
    {
        public UserSettingsInfo()
        {
            UserSettingsFileName = "userSettings.json";
            SchemaValue = @"{
  ""title"": ""Actions"",
  ""description"": ""User actions"",
  ""type"": ""object"",
  ""properties"": {
    ""Actions"": {
      ""type"": ""array"",
      ""minItems"": 1,
      ""maxItems"": 4,
      ""uniqueItems"": true,
      ""items"": {
        ""type"": ""object"",
        ""properties"": {
          ""Name"": {
            ""enum"": [
              ""GenerateConsolidatedReport"",
              ""GenerateLeaveReport"",
              ""CalculatePunchMovement"",
              ""InOutEntry""
            ]
          },
          ""Run"": {
            ""type"": ""boolean""
          },
          ""InputFolder"": {
            ""type"": ""string""
          }
        },
        ""allOf"": [
          {
            ""if"": {
              ""properties"": {
                ""Name"": {
                  ""const"": ""GenerateConsolidatedReport""
                }
              }
            },
            ""then"": {
              ""properties"": {
                ""MonthlyReportIdCol"": {
                  ""type"": ""number"",
                  ""default"": 4
                },
                ""MonthlyReportMonths"": {
                  ""type"": ""array"",
                  ""default"": [],
                  ""items"": {
                    ""type"": ""string""
                  }
                },
                ""PtrBookingMonthCol"": {
                  ""type"": ""number"",
                  ""default"": 4
                },
                ""PtrBookingMonths"": {
                  ""type"": ""array"",
                  ""default"": [],
                  ""items"": [
                    {
                      ""type"": [ ""number"", ""string"" ]
                    }
                  ],
                  ""additionalItems"": false
                },
                ""PtrEffortCols"": {
                  ""type"": ""array"",
                  ""default"": [],
                  ""items"": {
                    ""type"": ""number""
                  }
                },
                ""PtrProjectIdCol"": {
                  ""type"": ""number"",
                  ""default"": 4
                },
                ""PtrSheetName"": {
                  ""type"": ""string""
                }
              }
            }
          },
          {
            ""if"": {
              ""properties"": {
                ""Name"": {
                  ""const"": ""GenerateLeaveReport""
                }
              },
              ""required"": [
                ""Name""
              ]
            },
            ""then"": {
              ""properties"": {
                ""FinancialYear"": {
                  ""type"": ""string""
                }
              }
            }
          },
          {
            ""if"": {
              ""properties"": {
                ""Name"": {
                  ""const"": ""CalculatePunchMovement""
                }
              },
              ""required"": [
                ""Name""
              ]
            },
            ""then"": {
              ""properties"": {
                ""CutOff"": {
                  ""type"": ""string""
                }
              }
            }
          },
          {
            ""if"": {
              ""properties"": {
                ""Name"": {
                  ""const"": ""InOutEntry""
                }
              },
              ""required"": [
                ""Name""
              ]
            },
            ""then"": {
              ""properties"": {}
            }
          }
        ],
        ""required"": [
          ""Name"",
          ""Run"",
          ""InputFolder""
        ]
      }
    }
  },
  ""required"": [
    ""Actions""
  ]
}";
        }

        public string SchemaValue { readonly get; set; }
        public string UserSettingsFileName { readonly get; set; }
    }

    public class Action
    {
        [JsonPropertyName("Name")]
        public required string Name { get; set; } = string.Empty;

        [JsonPropertyName("Run")]
        public required bool Run { get; set; }

        [JsonPropertyName("InputFolder")]
        public required string InputFolder { get; set; } = string.Empty;

        [JsonPropertyName("MonthlyReportIdCol")]
        public int MonthlyReportIdCol { get; set; } = -1;

        [JsonPropertyName("MonthlyReportMonths")]
        public object[] MonthlyReportMonths { get; set; } = [];

        [JsonPropertyName("PtrBookingMonthCol")]
        public int PtrBookingMonthCol { get; set; } = -1;

        [JsonPropertyName("PtrBookingMonths")]
        public object[] PtrBookingMonths { get; set; } = [];

        [JsonPropertyName("PtrEffortCols")]
        public object[] PtrEffortCols { get; set; } = [];

        [JsonPropertyName("PtrProjectIdCol")]
        public int PtrProjectIdCol { get; set; } = -1;

        [JsonPropertyName("PtrSheetName")]
        public string PtrSheetName { get; set; } = string.Empty;

        [JsonPropertyName("FinancialYear")]
        public string FinancialYear { get; set; } = string.Empty;

        [JsonPropertyName("CutOff")]
        public string CutOff { get; set; } = string.Empty;
    }
}