Sub Ticker()

Dim Ticker As String
Dim i, Last_Row As Long
Dim Table_Row, Ticker_Count  As Integer
Dim Yearly_Change, Percent_Change, Total_Vol As Double
Dim ws As Worksheet
Dim myRange As Range


For Each ws In ThisWorkbook.Worksheets
ws.Activate

Columns("A:Q").AutoFit

'Formating

Set myRange = Range("J:J")
myRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
myRange.FormatConditions(1).Interior.Color = vbRed
myRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:=0
myRange.FormatConditions(2).Interior.Color = vbGreen

Range("A:Q").Select
Selection.Font.Bold = True



'Setting variabiables

Cells(2, 17).Value = Cells(2, 11).Value
Cells(3, 17).Value = Cells(2, 11).Value
Cells(4, 17).Value = Cells(2, 12).Value
Total_Vol = 0
Table_Row = 2
Last_Row = ActiveSheet.UsedRange.Rows.Count


'Table_1
Cells(1, 9) = "Ticker"
Cells(1, 12) = "Total Volume"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"

'Table_2
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(4, 15) = "Greatest total volume"
Cells(2, 15) = "Greatest percentage increase"
Cells(3, 15) = "Greatest percentage decrease"

'The code

For i = 2 To Last_Row



If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

row_count_ticker = Cells(i, 1).Rows.Count


    Ticker = Cells(i, 1).Value
    Cells(Table_Row, 9) = Ticker

    Yearly_Change = Cells(i, 6).Value - Cells(i - 250, 3).Value
    Cells(Table_Row, 10) = Yearly_Change

    Percent_Change = ((Cells(i, 6).Value - Cells(i - 250, 3).Value) / Cells(i - 250, 3).Value) * 100
    Cells(Table_Row, 11) = Percent_Change

    Total_Vol = Total_Vol + Cells(i, 7).Value
    Cells(Table_Row, 12) = Total_Vol





If Cells(Table_Row, 11).Value >= Cells(2, 17).Value Then
    Cells(2, 16) = Cells(Table_Row, 9).Value
    Cells(2, 17) = Cells(Table_Row, 11)
End If

If Cells(Table_Row, 11).Value <= Cells(3, 17).Value Then
    Cells(3, 16) = Cells(Table_Row, 9).Value
    Cells(3, 17) = Cells(Table_Row, 11).Value
End If


If Cells(Table_Row, 12) >= Cells(4, 17).Value Then
    Cells(4, 16) = Cells(Table_Row, 9).Value
    Cells(4, 17) = Cells(Table_Row, 12).Value
End If

Table_Row = Table_Row + 1
Total_Vol = 0


ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    Total_Vol = Total_Vol + Cells(i, 7).Value
End If

Next i
Next ws

End Sub
