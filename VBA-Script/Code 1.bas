Attribute VB_Name = "Module2014"
Sub stock_data()

Dim Ticker As String

Dim Total_Stock_Volume As Double

Total_Stock_Volume = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

For i = 2 To 705714

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Ticker = Cells(i, 1).Value

Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

Range("I" & Summary_Table_Row).Value = Ticker

Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

Summary_Table_Row = Summary_Table_Row + 1

Total_Stock_Volume = 0

Else

Total_Stock_Volume = Total_Stock_Volume = Cells(i, 7).Value

End If

Next i

End Sub
