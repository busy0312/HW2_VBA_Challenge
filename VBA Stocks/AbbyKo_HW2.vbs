 Sub alpha()

Dim ws As Worksheet
For Each ws In Worksheets

Dim summary As Integer
    summary = 2
Dim totalvol As Double
    totalvol = 0
Dim Yearly_open As Double
    Yearly_open = 0
Dim Yearly_close As Double
    Yearly_close = 0
Dim Yearly_change As Double
    


    

    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = " Percent_Change"
    ws.Range("L1").Value = "Total_StockVol"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    
    Yearly_open = ws.Cells(2, 3).Value


    For i = 2 To lastrow

        
  
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            totalvol = totalvol + ws.Cells(i, 7).Value
           ws.Range("I" & summary).Value = ws.Cells(i, 1).Value
            ws.Range("L" & summary).Value = totalvol
            
            
            Yearly_close = ws.Cells(i, 6).Value
            Yearly_change = Yearly_close - Yearly_open
             ws.Range("J" & summary).Value = Yearly_change
             
     If (Yearly_change > 0) Then
                 ws.Range("J" & summary).Interior.ColorIndex = 4
        ElseIf (Yearly_change <= 0) Then
                ws.Range("J" & summary).Interior.ColorIndex = 3

         End If
    
    If Yearly_open = 0 Then
            percentage = 0
    Else
          percentage = Yearly_change / Yearly_open
      
      End If
           
            ws.Range("K" & summary).Value = percentage
            ws.Range("K" & summary).NumberFormat = "0.00%"
           summary = summary + 1
           totalvol = 0
            
    
          Yearly_open = ws.Cells(i + 1, 3).Value
        
           
        Else
             totalvol = totalvol + ws.Cells(i, 7).Value
        
    
          
        End If
        
Next i



Dim greatest_increase As Double
greatest_increase = 0
Dim greatest_decrease As Double
greatest_decrease = 0

lastrow2 = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

For x = 2 To lastrow2

If ws.Cells(x, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow2)) Then
greatest_increase = greatest_increase + ws.Cells(x, 11).Value
ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
ws.Cells(2, 17).Value = greatest_increase
ws.Cells(2, 17).NumberFormat = "0.00%"

ElseIf ws.Cells(x, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow2)) Then
greatest_decrease = greatest_decrease + ws.Cells(x, 11).Value
ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
ws.Cells(3, 17).Value = greatest_decrease
ws.Cells(3, 17).NumberFormat = "0.00%"

End If



Next x


Dim greatest_vol As Double
greatest_vol = 0

lastrow3 = ws.Cells(Rows.Count, "L").End(xlUp).Row
For k = 2 To lastrow3

If ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow2)) Then
greatest_vol = greatest_vol + ws.Cells(k, 12).Value
ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
ws.Cells(4, 17).Value = greatest_vol
End If
Next k


   
Next ws

End Sub






















