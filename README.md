# VBA-Challenge
VBA coding language and visual studio was used to create this project.

## Usage/ Code 
Sub Ticker_Analysis()

    ' Declare Current as a worksheet object variable.
         Dim w As Worksheet
         For Each w In Worksheets


    w.Range("I1").Value = "ticker"
    w.Range("J1").Value = "yearly change"
    w.Range("K1").Value = "percent change"
    w.Range("L1").Value = "total stock volume"

    Dim Volume As Double
    Volume = 0
    
    
    Dim summary_counter As Double

    summary_counter = 2
    
    Dim opening_price_counter As Double
    opening_price_counter = 2
    
    Dim change As Double
    Dim percentage As Double
    
    
     
    
    
    
    
    RowCount = w.Cells(Rows.Count, "A").End(xlUp).Row
    'For Loop
    For i = 2 To RowCount
    
    
    If w.Cells(i, 1).Value <> w.Cells(i + 1, 1).Value Then
     Volume = Volume + w.Cells(i, 7).Value
     
     change = w.Cells(i, 6).Value - w.Cells(opening_price_counter, 3).Value
     percentage = change / w.Cells(opening_price_counter, 3).Value
     
     
        w.Range("I" & summary_counter).Value = w.Cells(i, 1).Value
         w.Range("J" & summary_counter).Value = change
         w.Range("J" & summary_counter).NumberFormat = "0.00"
          w.Range("K" & summary_counter).Value = percentage
           w.Range("K" & summary_counter).NumberFormat = "0.00%"
          
         If w.Range("J" & summary_counter).Value > 0 Then
         ' Set the Cell Colors to Red
  w.Range("J" & summary_counter).Interior.ColorIndex = 4
  
  
    ElseIf w.Range("J" & summary_counter).Value < 0 Then
         ' Set the Cell Colors to Red
  w.Range("J" & summary_counter).Interior.ColorIndex = 3
  
  End If
  
  
        
         w.Range("L" & summary_counter).Value = Volume
         
         
    summary_counter = summary_counter + 1
    Volume = 0
    
    change = 0
    
    opening_price_counter = i + 1
    
    
    
    
    Else
    Volume = Volume + w.Cells(i, 7).Value
    
    
    End If
    
    
    
    
    
    Next i
    
    


 Next w

    
End Sub



#### Outputs 
-Outputs for tickers, yearly change, percent change, and total stock volume were recorded for the years 2018,2019, and 2020.
- With outputs in consideration the greatest % increase, Greatest % Decrease, and Greatest total volume was measured for each Ticker. 

#### Importance of this measure 
Seeing how all aspects effect the total volume increase or decrease over a course of three years. 
