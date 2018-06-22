Sub AnnualVolume()

 ' Set an initial variable for holding the ticker symbol
 Dim Ticker_Symbol As String

' Set an initial variable for holding the total volume per ticker symbol
 Dim Annual_Volume As Double
 Annual_Volume = 0

 Dim LastRow As Long

 LastRow = Cells(Rows.Count, "A").End(xlUp).Row

 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2
 
' Loop through all ticker transactions
 For i = 2 To LastRow

   ' Check if we are still within the same ticker symbol, if we are not...
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

     ' Set the Ticker symbol
     Ticker_Symbol = Cells(i, 1).Value

     ' Add to the Annual Volume Total
     Annual_Volume = Annual_Volume + Cells(i, 7).Value

     ' Print the Ticker Symbol in the Summary Table
     Range("i" & Summary_Table_Row).Value = Ticker_Symbol

     ' Print the Annual Volume Amount to the Summary Table
     Range("j" & Summary_Table_Row).Value = Annual_Volume

     ' Add one to the summary table row
     Summary_Table_Row = Summary_Table_Row + 1 

     ' Reset the Annual Volume
     Annual_Volume = 0
   ' If the cell immediately following a row is the same ticker...

   Else
     ' Add to the Annual Volume
     Annual_Volume = Annual_Volume + Cells(i, 7).Value

   End If

 Next i

 ' Insert headers for new columns
  Range("I1").value = "Ticker"
  Range("J1").value = "Annual Stock Volume"
     ' Autofit to display data
      Worksheets("2016").Columns("I:J").AutoFit
    With Range("I:J")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With

End Sub