Sub MYStockData()

'Iterates through every worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate


'Creates title fields for columns and greatest value information
'Sets headers in row 1
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Sets Greatest column fields
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Determine number of active rows and sets result to variable
Dim RowCount As Long
RowCount = ActiveSheet.UsedRange.Rows.Count - 1
'For debug/testing purposes only. Writes RowCount to unused column.
'Cells(4, 13).Value = RowCount


'Establish and set values for variables used in main loop
Dim FindTicker As String
Dim TickerColumnRow As Long
Dim VolCounter As LongLong
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double

SummaryRow = 2
VolCounter = 0
OpenPrice = 1
ClosePrice = 1
YearlyChange = 0
PercentChange = 0

'Iterate through all columns
For i = 2 To RowCount + 1
'If current and next ticker aren't the same then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      'Get the ticker symbol and set equal to FindTicker variable
       FindTicker = Cells(i, 1).Value
       'Write ticker to column 9 ticker
       Cells(SummaryRow, 9).Value = FindTicker
       'Adds volume
       VolCounter = VolCounter + Cells(i, 7).Value
       'Writes volume total
       Cells(SummaryRow, 12).Value = VolCounter
       'Resets volume counter
       VolCounter = 0
       'Sets close price
       ClosePrice = Cells(i, 6).Value
       Cells(SummaryRow, 20).Value = ClosePrice
       'Calc and sets YearlyChange
       YearlyChange = ClosePrice - OpenPrice
       Cells(SummaryRow, 10).Value = YearlyChange
       'Calcs Percent Change
       PercentChange = (ClosePrice / OpenPrice) - 1
       'Formats to percentage
       Cells(SummaryRow, 11).Value = PercentChange
       Cells(SummaryRow, 11).NumberFormat = "0.00%"
       'Conditional format percent change column. Green for +, Red for -, and No fill for 0
       If Cells(SummaryRow, 11).Value > 0 Then
           Cells(SummaryRow, 11).Interior.ColorIndex = 4
           ElseIf Cells(SummaryRow, 11).Value < 0 Then
               Cells(SummaryRow, 11).Interior.ColorIndex = 3
           Else
              Cells(SummaryRow, 11).Interior.ColorIndex = 0
       End If
       'Conditional format Yearly Change column Green for +, Red for -, and No fill for 0
       If Cells(SummaryRow, 10).Value > 0 Then
           Cells(SummaryRow, 10).Interior.ColorIndex = 4
           ElseIf Cells(SummaryRow, 10).Value < 0 Then
               Cells(SummaryRow, 10).Interior.ColorIndex = 3
           Else
               Cells(SummaryRow, 10).Interior.ColorIndex = 0
       End If
        
      'Increments summary row variable
       SummaryRow = SummaryRow + 1
    Else
       'Adds last row volume to total
       VolCounter = VolCounter + Cells(i, 7).Value
       'Check for OpenPrice and write value
         If Len(Trim(Cells(SummaryRow, 19))) = 0 Then
           OpenPrice = Cells(i, 3).Value
           Cells(SummaryRow, 19) = OpenPrice
       'Sets ClosePrice
       ClosePrice = Cells(i + 1, 6).Value
       Cells(SummaryRow, 20).Value = ClosePrice
         End If
     End If
           
Next i

'Determines Summary Row Count
Dim SummaryRowCount As Long
SummaryRowCount = Cells(Rows.Count, 9).End(xlUp).Row

'For Debug/Testing only. Writes SummaryRowCount to unused column.
'Cells(5, 13).Value = SummaryRowCount

'Remove temp columns
ActiveSheet.Columns(20).ClearContents
ActiveSheet.Columns(19).ClearContents


'Find greatest % increase
Dim maxValue As Double
Dim maxTicker As String

maxValue = 0
maxTicker = " "

'Loops through column to determine, set, and write maxValue variable
For i = 2 To SummaryRowCount
    If Cells(i, 11).Value > maxValue Then
        maxValue = Cells(i, 11).Value
        Cells(2, 17).Value = maxValue
        Cells(2, 17).NumberFormat = "0.00%"
        maxTicker = Cells(i, 9).Value
        Cells(2, 16).Value = maxTicker
    End If
Next i

'Find greatest % decrease
Dim minValue As Double
Dim minTicker As String
minValue = 0
minTicker = " "

'Loops through column to determine, set, and write minValue variable
For i = 2 To SummaryRowCount
    If Cells(i, 11).Value < minValue Then
        minValue = Cells(i, 11).Value
        Cells(3, 17).Value = minValue
        Cells(3, 17).NumberFormat = "0.00%"
        minTicker = Cells(i, 9).Value
        Cells(3, 16).Value = minTicker
    End If
Next i

''Loops through column to determine, set, and write greatest total volume
Dim maxVol As Double
Dim volTicker As String
volTicker = " "
maxVol = 0

For i = 2 To SummaryRowCount
    If Cells(i, 12).Value > maxVol Then
        maxVol = Cells(i, 12).Value
        Cells(4, 17).Value = maxVol
        volTicker = Cells(i, 9).Value
        Cells(4, 16).Value = volTicker
    End If
Next i

'Upon completion iterations to next worksheet, if exists, and starts code again.
Next ws

End Sub
