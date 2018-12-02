Sub worksheetloop()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

Dim Ticker As String
Dim Total_stockvolume As Double
Total_stockvolume = 0

Dim yearly_change As Double
Dim percent_change As Double
Dim open_price As Double
Dim close_price As Double


Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"
open_price = Cells(2, 3).Value

For i = 2 To Lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
close_price = Cells(i, 6).Value
yearly_change = close_price - open_price
If (open_price = 0 And close_price = 0) Then
percent_change = 0
ElseIf (open_price = 0 And close_price <> 0) Then
percent_change = 1
Else: percent_change = yearly_change / open_price
Total_stockvolume = Total_stockvolume + ws.Cells(i, 7).Value
ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Range("L" & Summary_Table_Row).Value = Total_stockvolume
ws.Range("J" & Summary_Table_Row).Value = yearly_change
ws.Range("K" & Summary_Table_Row).Value = percent_change
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
End If
If percent_change > 0 Then
ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf percent_change < 0 Then
ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
End If

Summary_Table_Row = Summary_Table_Row + 1
Total_stockvolume = 0
open_price = Cells(i + 1, 3).Value

Else
Total_stockvolume = Total_stockvolume + ws.Cells(i, 7).Value
End If
Next i

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

PClastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
For j = 2 To PClastrow
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
If Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & PClastrow)) Then
Cells(2, 17).Value = Cells(j, 11).Value
Cells(2, 17).NumberFormat = "0.00%"
Cells(2, 16).Value = Cells(j, 9).Value
ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & PClastrow)) Then
Cells(3, 17).Value = Cells(j, 11).Value
Cells(3, 17).NumberFormat = "0.00%"
Cells(3, 16).Value = Cells(j, 9).Value
ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & PClastrow)) Then
Cells(4, 17).Value = Cells(j, 12).Value
Cells(4, 16).Value = Cells(j, 9).Value
End If
Next j

Next ws
MsgBox ("Done! Your resluts are here!")
End Sub



