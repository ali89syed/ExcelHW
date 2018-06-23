Sub stock_data()

'set the initial variable for holding the stock name'
Dim Ticker_Name As String


'total stock type'

Dim Stock_total As Double

' Keep track of the location for each stock in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
  
'loop through all stocks'

For i = 2 To 79771

'check to see similar stocks names'

If Cells(i + 1, 1).Value <> Cells(i, 1) Then


Ticker_Name = Cells(i, 1).Value



 ' Add to the stock total
      Stock_total = Stock_total + Cells(i, 3).Value

' Print the stock name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the stock amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = Stock_total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the stock Total
      Stock_total = 0


    Else

      Stock_total = Stock_total + Cells(i, 3).Value

    End If
  
Next i

End Sub