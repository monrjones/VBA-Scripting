Sub tickerPull()
Dim ws As Worksheet

    ' Loop through all of the worksheets in the active workbook.
For Each ws In ActiveWorkbook.Worksheets
  Dim tickerName As String
  Dim volumeTotal As Double
        volumeTotal = 0
 
Dim tickerNameRow As Integer
 tickerNameRow = 2

 lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



  For i = 2 To lastRow
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
      tickerName = ws.Cells(i, 1).Value

      
      volumeTotal = volumeTotal + ws.Cells(i, 7).Value

      
      ws.Range("I" & tickerNameRow).Value = tickerName

     
      ws.Range("J" & tickerNameRow).Value = volumeTotal

      
      tickerNameRow = tickerNameRow + 1
      
      
      volumeToral = 0

    
    Else

      
      volumeTotal = volumeTotal + ws.Cells(i, 7).Value

    End If

  Next i
Next ws

End Sub

