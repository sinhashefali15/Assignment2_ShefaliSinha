Sub multiple_year_stock()

'To simultaneously run all the three sheet
For Each ws in Worksheets

  'initializing variable to hold Ticker
  Dim Ticker As String

  

  'initializing variable for total volume
  Dim Volume_Total As Double
  Volume_Total = 0

  
  Dim Summary As Integer
  Summary = 2
  
  'Get the last row value
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all the Tickers
  For i = 2 To LastRow

    'if condition to check if it is the same ticker, if not execute below
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'Get the new Ticker value
      Ticker = ws.Cells(i, 1).Value

      'Sum of the total volume
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

    

      'Printing the Ticker 
      ws.Range("L" & Summary).Value = Ticker

      'Printing the total volume
      ws.Range("M" & Summary).Value = Volume_Total

      'Add one to the summary table row
      Summary = Summary + 1
      
      'Reset total volume
      Volume_Total = 0

    ' If same ticker
    Else

      ' Add to the total volume
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

    End If

  Next i

Next ws 

End Sub
