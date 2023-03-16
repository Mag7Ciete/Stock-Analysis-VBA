Attribute VB_Name = "Module1"

Sub StockMTicker()


'First  we  establish variables

Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Volume As Long
Dim TotalXVol As Double
Dim ResultsRange As Integer

Dim LastRow As Long
Dim start As Long


'Making sure instructions apply to each worksheet -Year

For Each ws In ActiveWorkbook.Worksheets

'Setting the values

TotalXVol = 0
ResultsRange = 2
start = 2


'Creating Labes for columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly_Change"
ws.Range("K1").Value = "Percent_Change"
ws.Range("L1").Value = "Total_Stock_Volume"



'Bonus Section
'Labels for heather and totals

ws.Range("O2").Value = "Greatest% Increase"
ws.Range("O3").Value = "Greatest% Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Setting last Row

 LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
 
 'Loops &  Loops  and Set values for loops & makng sure the loop considers the very last row
    
 For e = 2 To LastRow
 
    
    If e = LastRow + 1 Then
    
        End If
        
 'Finding ticker same  ang grouping by name also calculating : total volume, diff between open and close price, percentage per loss or gain.
 
If ws.Cells(e + 1, 1).Value <> ws.Cells(e, 1).Value Then

Ticker = ws.Cells(e, 1).Value


TotalXVol = TotalXVol + ws.Cells(e, 7).Value

OpenPrice = ws.Cells(start, 3).Value

ClosePrice = ws.Cells(e, 6).Value

Yearly_Change = ClosePrice - OpenPrice

Percent_Change = Yearly_Change / OpenPrice

start = e + 1

ws.Cells(ResultsRange, 9).Value = Ticker
ws.Cells(ResultsRange, 10).Value = Yearly_Change
ws.Cells(ResultsRange, 11).Value = Percent_Change
ws.Cells(ResultsRange, 12).Value = TotalXVol

ws.Cells(ResultsRange, 11).NumberFormat = "0.00%"

If Yearly_Change >= 0 Then

ws.Cells(ResultsRange, 10).Interior.ColorIndex = 4

Else
ws.Cells(ResultsRange, 10).Interior.ColorIndex = 3

End If

ResultsRange = ResultsRange + 1

TotalXVol = 0

Else

TotalXVol = TotalXVol + ws.Cells(e, 7).Value

End If

Next e

Ticker = " "
TotalXVol = 0

Next ws





End Sub

