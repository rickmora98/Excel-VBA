VBA-challenge

This repository contains my solutions to Data Analytics Bootcamp Homework Assignment #2.

Files:

"StockAnalyzer.vbs" is the VBA module to use for this assignment  
"2014 Stock Analysis.JPG" is a screenshot of worksheet 2014 after running script
"2015 Stock Analysis.JPG" is a screenshot of worksheet 2015 after running script
"2016 Stock Analysis.JPG" is a screenshot of worksheet 2016 after running script

"StockAnalyzer.vbs" contains 2 subroutines:

SingleYearStockAnalyzer(): 
	works only on the worksheet that is active
MultiYearStockAnalyzer(): 
	loops through all sheets in the active workbook and calls SingleYearStockAnalyzer()

Note:

Stock ticker "PLNT" on year 2015 had values of 0 on more than half of it's rows.
To avoid any errors, the subroutine skips any rows that contain only zeros.



Submitted by Ricardo G. Mora, Jr.  09/22/2021	