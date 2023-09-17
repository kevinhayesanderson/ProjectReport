# Project Report Application Information

## 1\. User Setting information

<font size= "4">
The user settings are stored in the userSettings.json file.<br/>
Here’s an example userSettings.json for reference:

    {
        "Folder": "C:\\Users\\Id\\Reports\\InputDir",
        "MonthlyReportMonths": [ "Jun-23", "Jul-23", "Aug-23" ],
        "PtrProjectIdCol": 2,
        "PtrBookingMonthCol": 20,
        "PtrBookingMonths": [ "6|June", 7, "8|Aug_22" ],
        "PtrEffortCols": [ 18 ],
        "PtrSheetName": "PTR_23",
        "GenerateLeaveReport": false,
        "FinancialYear": "23-24"
    }

***Folder:***
Path to reports and output path of the application.

***MonthlyReportMonths:***
List of sheet names of monthly reports to read.
>***options***
> `[]` - All sheets will be read from the monthly report.  
> `["Jan-22"]` - Only sheet with name “Jan-22” will be read.  
> `["Jan-22","Feb-22"]` - Sheets with name “Jan-22” & “Feb-22” will be read.

***PtrProjectIdCol:***
Column number indicating project ID column in Ptr report.

***PtrBookingMonthCol:***
Column number indicating booking month column in Ptr report.

***PtrBookingMonths:***
List of values for booking month column in Ptr report.
> ***options***  
> `[]` - All booking months project IDs will be read.  
> `[3]` - Only booking month 3 project IDs will be read.  
> `[3,4]` - Project IDs for booking months 3 and 4 will be read.  
> `["1|Jan-22|Jan|Jan_23"]` - Only booking month 1 with formats “Jan-22”, "Jan" & "Jan_23" project IDs will be read.  
> `["5|May-22|May", "6|June_2023"]` - Project IDs for booking months 5 or “May-22”, "May" and 6 or "June_2023" will be read.

***PtrEffortCols:***
List of column numbers for effort column in Ptr report.
> ***options***  
> `[22]` - Only effort in column 22 will be read.  
> `[22,23]` - Summation of efforts in columns 22 and 23 will be read.

***PtrSheetName:***
Sheet name to read values in the Ptr report.

***GenerateLeaveReport:***
Generation of leave report.
> ***options***  
> `true` - The application will generate a leave report and will not generate consolidated and inter reports.  
> `false` - The application will not generate a leave report and will generate consolidated and inter reports.

***FinancialYear:***
The single financial year for which the leave report will be generated.
> ***format***  
> `"20-21"` - FY 2020-2021.  
> `"21-22"` - FY 2021-2022.

## 2\. Monthly report constraints

1. The file name should have “Monthly_Report” in the file name to be considered as a Monthly report.
2. All sheets should have a default layout and default format for correct data capture.
3. Available time-Actual will be read from the last column in row 13.
4. Leave will be read from the last column in row 14.
5. Project rows will be read from row 16.
6. Project Ids should start with “ACS\_” or “ACS.” or “Acs\_” or “Acs.” or “acs\_” or “acs.”.

## 3\. PTR constraints

1. The file name should have “ACS_PTR” present in the file name to be considered as a PTR report.
2. Effort should either be in the Excel format of `number` or `[h]:mm:ss`.

## 4\. General constraints

1. Year-wise PTR will be read faster as the filter iterates through all the rows in the sheet.
</font>
