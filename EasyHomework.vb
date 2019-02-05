Sub Homework()

'Create Ticker and Total Stock Volume variables
Dim Ticker As String
Dim Total_Stock_Volume As Double

'Create a variable for the Last Row in the ws
Dim Last_Row As Double
Last_Row = Cells(Rows.Count, "A").End(xlUp).Row

 ' Keep track of the location for each Ticker in a separate table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'Set I and J Column names to "Ticker" and "Total Stock Volume", respectively
Range("$I$1").Value = "Ticker"
Range("$J$1").Value = "Total Stock Volume"

' Loop through all ticker symbols
  For i = 2 To Last_Row

    ' Check if next row is the same ticker symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = Cells(i, 1).Value

      ' Add to the Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print the Ticker symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Total_Stock_Volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Increment the Summary_Table_Row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total_Stock_Volume for the next Ticker
      Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i

End Sub
