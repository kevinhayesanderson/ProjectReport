## Project Report Application Information
### Application Description:
&emsp; The Project Report application is a windows console application used to create consolidated reports from Monthly and PTR reports.
### Application Source code:
&emsp;[https://github.com/kevinhayesanderson/ProjectReport](https://github.com/kevinhayesanderson/ProjectReport)

### Application Owner Info: 
&emsp;MohanRaj L.V.P 
&emsp;[mohanraj.lvp@ambigai.net](mailto:mohanraj.lvp@ambigai.net)
 
### Developer Info:
&emsp;Kevin Hayes Anderson
&emsp;[kevin.hayes@ambigai.net](mailto:kevin.hayes@ambigai.net)
 
 
### User Setting information:
The user settings are stored in the userSettings.json file.<br/>
It contains an array of actions.
Here's an example of userSettings.json:
```json
{
    "Actions": [
        {
            "Name": "GenerateConsolidatedReport",
            "Run": true,
            "InputFolder": "E:\\ProjectReport\\Requirements\\Second_requirement",
            "MonthlyReportIdCol": 0,
            "MonthlyReportMonths": [],
            "PtrBookingMonthCol": 0,
            "PtrBookingMonths": [],
            "PtrEffortCols": [],
            "PtrProjectIdCol": 0,
            "PtrSheetName": ""
        },
        {
            "Name": "GenerateLeaveReport",
            "Run": true,
            "InputFolder": "E:\\ProjectReport\\Requirements\\Second_requirement",
            "FinancialYear": ""
        },
        {
            "Name": "CalculatePunchMovement",
            "Run": true,
            "InputFolder": "E:\\ProjectReport\\Requirements\\Second_requirement"
        }
    ]
}
```
### Actions:
Each action contains three common properties: Name, Run & InputFolder.
 
***Name:***
Name of the action. 
> :warning: Do not edit this value.
 
***Run:***
Boolean property which indicates whether the action should be executed. 
 
***InputFolder:***
Path to input reports. 
 
### GenerateConsolidatedReport:
This action generates a consolidated report.
 
***MonthlyReportIdCol:***
Column number indicating project ID column in a monthly report.
The default value is 3.
 
***MonthlyReportMonths:***
List of sheet names of monthly reports to read.
>***options***
> `[]` - All sheets will be read from the monthly report.  
> `["Jan-22"]` - Only sheet with name "Jan-22" will be read.  
> `["Jan-22","Feb-22"]` - Sheets with name "Jan-22" & "Feb-22" will be read.
 
***PtrBookingMonthCol:***
Column number indicating booking month column in Ptr report.
 
***PtrBookingMonths:***
List of values for booking month column in Ptr report.
> ***options***  
> `[]` - All booking months project IDs will be read.  
> `[3]` - Only booking month 3 project IDs will be read.  
> `[3,4]` - Project IDs for booking months 3 and 4 will be read.  
> `["1|Jan-22|Jan|Jan_23"]` - Only booking month 1 with formats "Jan-22", "Jan" & "Jan_23" project IDs will be read.  
> `["5|May-22|May", "6|June_2023"]` - Project IDs for booking months 5 or "May-22", "May" and 6 or "June_2023" will be read.
 
***PtrEffortCols:***
List of column numbers for effort column in Ptr report.
> ***options***  
> `[22]` - Only effort in column 22 will be read.  
> `[22,23]` - Summation of efforts in columns 22 and 23 will be read.
 
***PtrProjectIdCol:***
Column number indicating project ID column in Ptr report.
 
***PtrSheetName:***
Sheet name to read values in the Ptr report.
 
### GenerateLeaveReport:
This action generates a leave report.
 
***FinancialYear:***
The single financial year for which the leave report will be generated.
> ***format***  
> `"20-21"` - FY 2020-2021.  
> `"21-22"` - FY 2021-2022.
 
### CalculatePunchMovement:
This action generates a time report.
 
### Monthly report constraints:
 
1. The file name should have "Monthly_Report" in the file name to be considered as a Monthly report.
2. All sheets should have a default layout and default format for correct data capture.
3. Available time-Actual will be read from the last column in row 13.
4. Leave will be read from the last column in row 14.
5. Project rows will be read from row 16.
6. Project Ids should start with "ACS\_" or "ACS." or "Acs\_" or "Acs." or "acs\_" or "acs.".
 
### PTR constraints:
 
1. The file name should have "ACS_PTR" present in the file name to be considered as a PTR report.
2. Effort should either be in the Excel format of `number` or `[h]:mm:ss`.
 
### General constraints:
 
1. Year-wise PTR will be read faster as the filter iterates through all the rows in the sheet.
<br/><br/>
**<p style="text-align: center;">&copy; 2023 [Ambigai Consultancy Services GmbH](https://www.ambigai.net/)</p>**
