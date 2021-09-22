Attribute VB_Name = "Module1"
Option Explicit '(requires that all variables be declared)



Sub MultiYearStockAnalyzer()

'   Subroutine: MultiYearStockAnalyzer
'   Created by: Ricardo G. Mora, Jr. 09/20/2021
'   Updated by:
'
'   Purpose: To output a summary table on each sheet within a workbook that contains for each stock
'   it's change in price from beginning of year to end of year, the percent change in price, and
'   cumulative volume.  This subroutine also adds a second table to each sheet that indicates the
'   stocks with the greatest percent increase, greatest percent decrease, and greatest cumulative volume.
'
'   Paramaters: None
'   Additional Subroutines/Functions required: SingleYearStockAnalyzer


'   Declare all variables:
'
    Dim StockDataSheet As Worksheet

'   Loop through each worksheet in the Active Workbook and
'   invoke subroutine SingleYearStockAnalyzer:
'
    For Each StockDataSheet In Worksheets
        StockDataSheet.Select
        Call SingleYearStockAnalyzer
    Next StockDataSheet

End Sub



Sub SingleYearStockAnalyzer()

'   Subroutine: SingleYearStockAnalyzer
'   Created by: Ricardo G. Mora, Jr.  09/20/2021
'   Updated by: Ricardo G. Mora, Jr.  09/22/2021
'
'   Purpose: To output a summary table on the Active Sheet that contains for each stock
'   it's change in price from beginning of year to end of year, the percent change in price, and
'   cumulative volume.  This subroutine also adds a second table to the sheet that indicates the
'   stocks with the greatest percent increase, greatest percent decrease, and greatest cumulative volume.
'
'   Parameters: None
'   Additional Subroutines/Functions required: None
'   Input Data Format: Normalized 7 column table with top row as headers:
'       Column1: <ticker> as String (stock ticker symbol)
'       Column2: <date> as Integer with format "yyyymmdd" (stock trading date)
'       Column3: <open> as Double (stock opening bell price)
'       Column4: <high> as Double (stock highest price during day)
'       Column5: <low> as Double (stock lowest price during day)
'       Column6: <close> as Double (stock closing bell price)
'       Column7: <vol> as Integer (quantity of shares traded during day)
'
'   Note: PRIOR TO RUNNING THIS SUBROUTINE, THE INPUT DATA IN THE WORKSHEET MUST BE CLEAN
'   AND SORTED BY TICKER AND DATE.


'   Declare all variables:
'
    Dim LastRow As Long
    Dim RowNum As Long
    Dim OutputRow As Long
    Dim StockSymbol As String
    Dim StockTotalVolume As Variant '(declared as Variant since values exceeded the limits for Long)
    Dim StockStartPrice As Double
    Dim StockEndPrice As Double
    Dim StockPriceChange As Double
    Dim StockPercentChange As Double
    Dim GreatestIncreaseValue As Double
    Dim GreatestDecreaseValue As Double
    Dim GreatestVolumeValue As Variant '(declared as Variant since values exceeded the limits for Long)
    Dim GreatestIncreaseStock As String
    Dim GreatestDecreaseStock As String
    Dim GreatestVolumeStock As String
    
'   Create column and row headers for output data:
'
    Range("I1:L1").Value = Split("Ticker/Yearly Change/Percent Change/Total Stock Volume", "/")
    Range("P1:Q1").Value = Split("Ticker/Value", "/")
    Range("O2:O4").Value = WorksheetFunction.Transpose(Split("Greatest % Increase/Greatest % Decrease/Greatest Total Volume", "/"))
    Range("J1:L1,Q1").HorizontalAlignment = xlRight
    
'   Initialize variables:
'
    StockSymbol = Range("A2").Value
    StockStartPrice = Range("C2").Value
    StockTotalVolume = 0
    OutputRow = 2
    GreatestIncreaseValue = 0
    GreatestDecreaseValue = 0
    GreatestVolumeValue = 0
    
'   Loop through all rows containing data:
'
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For RowNum = 2 To LastRow
    
'       Update the running stock volume for the current stock:
'
        StockTotalVolume = StockTotalVolume + Range("G" & RowNum).Value
        
'       If the stock ticker symbol on the row following this one changes...
'
        If Range("A" & RowNum + 1).Value <> StockSymbol Then

'           Calculate stock price change:
'
            StockEndPrice = Range("F" & RowNum).Value
            StockPriceChange = StockEndPrice - StockStartPrice
            
'           Calculate stock percent change without causing a divide by zero error:
'
            If StockStartPrice = 0 Then
                StockPercentChange = 0
            Else
                StockPercentChange = StockPriceChange / StockStartPrice
            End If
            
'           Output ticker symbol, yearly price change and yearly percent change for the stock onto the sheet:
'
            Range("I" & OutputRow).Value = StockSymbol
            Range("J" & OutputRow).Value = Format(StockPriceChange, "Fixed")
            Range("K" & OutputRow).Value = Format(StockPercentChange, "Percent")
            Range("L" & OutputRow).Value = Format(StockTotalVolume, "#,###")
            
'           Apply color formatting to the percent change cell to indicate positive or negative change:
'
            If StockPriceChange >= 0 Then
                Range("J" & OutputRow).Interior.ColorIndex = 4  ' (green)
            Else
                Range("J" & OutputRow).Interior.ColorIndex = 3  ' (red)
            End If
            
'           Update the greatest increase, greatest decrease, and greatest volume variables if neccessary:
'
            If StockPercentChange > GreatestIncreaseValue Then
                GreatestIncreaseStock = StockSymbol
                GreatestIncreaseValue = StockPercentChange
            ElseIf StockPercentChange < GreatestDecreaseValue Then
                GreatestDecreaseStock = StockSymbol
                GreatestDecreaseValue = StockPercentChange
            End If
            If StockTotalVolume > GreatestVolumeValue Then
                GreatestVolumeStock = StockSymbol
                GreatestVolumeValue = StockTotalVolume
            End If
            
'           Initialize variables for the next stock and continue looping:
'
            StockSymbol = Range("A" & RowNum + 1).Value
            StockStartPrice = Range("C" & RowNum + 1).Value
            StockTotalVolume = 0
            OutputRow = OutputRow + 1
            
'       But if the stock ticker symbol on the row following this one doesn't change...
'       And the stock's opening price is 0, use the opening price on the following row
'       for the next run through this loop.  This is in case a stock doesn't have any
'       activity until mid year (as is the case with PLNT in 2015).
'
        ElseIf StockStartPrice = 0 Then
            StockStartPrice = Range("C" & RowNum + 1)
        End If
        
    Next RowNum
    
'   Output the stocks with the greatest percent increase, greatest percent decrease,
'   and greatest volume to the Active Sheet:
'
    Range("P2").Value = GreatestIncreaseStock
    Range("Q2").Value = Format(GreatestIncreaseValue, "Percent")
    Range("P3").Value = GreatestDecreaseStock
    Range("Q3").Value = Format(GreatestDecreaseValue, "Percent")
    Range("P4").Value = GreatestVolumeStock
    Range("Q4").Value = Format(GreatestVolumeValue, "#,###")
    
'   Make all output visible within their columns:
'
    Columns("I:L").EntireColumn.AutoFit
    Columns("O:Q").EntireColumn.AutoFit
    
End Sub
