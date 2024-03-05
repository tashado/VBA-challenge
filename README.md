**Name**
  Module 2 Challenge - VBA Scripting - Multiple Year Stock Data

**Description**
  Using the provided file from Monah Bootcamp, Module 2 Challenge, this code provides for each ticker:
  - yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
  - percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
  - total stock volume of the stock

  Also provides the ticker with:
  - the greatest percentage increase
  - the greatest percentage decrease
  - the greatest total volume

**Acknowledgment**
  Line 27 and 28 of the VBS.file to find the last row of the data set was taken from  Monash Bootcamp, file "census_data_2016-2019_pt1_solution.xlsm".
  Code is as below:
  
      Dim last_row As Long
      
      last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

