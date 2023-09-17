## Project Report Application Information

##### 1\. User Setting information:

The user settings are stored in **_userSettings.json_** file.  
Here’s a example **_userSettings.json_** for reference.

    {
        "Folder": "C:\\Users\\90258\\Downloads\\Reports",
        "MonthlyReportMonths": [ "Aug-22" ],
        "PtrProjectIdCol": 2,
        "PtrBookingMonthCol": 20,
        "PtrBookingMonths": [ "6|Jun-22", 7, "Aug-22" ],
        "PtrEffortCols": [ 18 ],
        "PtrSheetName": "PTR2",
        "GenerateLeaveReport": false,
        "FinancialYear": "21-22"
    }

Folder

Path to reports and output path of the application.

MonthlyReportMonths

List of sheet names of monthly report to read.

> **_options_**  
> `[]` - All sheet will be read from the monthly report.  
> `["Jan-22"]` - Only sheet with name “Jan-22” will be read.  
> `["Jan-22","Feb-22"]` - Sheets with name “Jan-22” & “Feb-22” will be read.

PtrProjectIdCol

Column number indicating project id column in ptr report.

PtrBookingMonthCol

Column number indicating booking month column in ptr report.

PtrBookingMonths

List of values for booking month column in ptr report.

> **_options_**  
> `[]` - All booking months project ids will be read.  
> `[3]` - Only booking month 3 project ids will be read.  
> `[3,4]` - Project ids for booking months 3 and 4 will be read.  
> `["Jan-22"]` - Only booking month “Jan-22” project ids will be read.  
> `["Jan-22", 2]` - Project ids for booking months “Jan-22” and 2 will be read.  
> `["5|May-22"]` - Project ids for booking months 5 or “May-22” will be read.

PtrEffortCols

List of column numbers for effort column in ptr report.

> **_options_**  
> `[22]` - Only effort in column 22 will be read.  
> `[22,23]` - Summation of efforts in column 22 and 23 will be read.

PtrSheetName

Sheet name to read values in prt report.

GenerateLeaveReport

Generation of leave report.

> **_options_**  
> `true` - Application will generate leave report and will not generate consolidated and inter reports.  
> `false` - Application will not generate leave report and will generate consolidated and inter reports.

FinancialYear

Single financial year for which the leave report will be generated.

> **_format_**  
> `"20-21"` - FY 2020-2021.  
> `"21-22"` - FY 2021-2022.

##### 2\. Monthly report constraints:

1.  File name should have “Monthly_Report” in the file name to be considered as Monthly report.
2.  All sheets should have default layout and default format for correct data capture.
3.  Name and Id for employee will be taken from sheet 1 of the monthly report.
4.  Project rows will be read from row 15.
5.  Project Ids should start with “ACS\_” or “ACS.” or “Acs\_” or “Acs.” or “acs\_” or “acs.”.

##### 3\. PTR report constraints:

1.  File name should have “ACS_PTR” present in the file name to be considered as PTR report.
2.  Effort should either be in the excel format of `number` or `[h]:mm:ss`.
3.  Booking month should either be in the excel format of `number` or `mmm-yy`.

##### 4\. General constraints:

1.  Year wise PTR will be read faster as the filter iterates through all the rows in the sheet.
