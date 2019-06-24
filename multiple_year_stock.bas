Attribute VB_Name = "Module1"
Sub stockmarketdata()

    For Each ws In Worksheets
        Dim ticker As String
        Dim volume As Double
        Dim tickertable As Integer
        
               
        volume = 0
        tickertable = 2
        
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            'MsgBox (ticker)
            
            
            ws.Range("i" & tickertable) = ticker
            ws.Range("j" & tickertable) = volume
            
       volume = volume + ws.Cells(i, 7).Value
       ws.Range("i1") = "Ticker"
       ws.Range("j1") = "Total Stock Volume"
         
       tickertable = tickertable + 1
       volume = 0
       
       Else
       
      volume = volume + ws.Cells(i, 7).Value
        
       End If

Next i
Next ws


End Sub
