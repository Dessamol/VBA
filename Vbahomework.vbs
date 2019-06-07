Attribute VB_Name = "Module1"
Sub Stock_looper():


Dim ws As Worksheet


For Each ws In ThisWorkbook.Worksheets

ws.Activate

Dim stock_total As Double

stock_total = 1

Dim summary_row As Double

summary_row = 2


For I = 2 To 70926
stock_total = stock_total + Cells(I, 7).Value

    If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
    ticker_name = Cells(I, 1).Value
    Cells(summary_row, 9).Value = ticker_name
     Cells(summary_row, 10).Value = stock_total
     stock_total = 0
     summary_row = summary_row + 1
      Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
  
   End If
   Next I
   
   Next ws




End Sub




