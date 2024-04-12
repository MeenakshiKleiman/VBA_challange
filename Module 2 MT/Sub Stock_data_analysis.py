Sub Stock_data_analysis()

    Dim Ticker As String
    Dim Total As Double
    Dim percentchange As Double
    Dim i As Long
    Dim j As Integer
    Dim PreviousStockPrice As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim yearlychange As Double
    
   'For looping through each worksheet
   
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    
    ' Summary Table column title
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    
    
    ¡ = 0
    Total = 0
    Start = 2
    PreviousStockPrice = 2

    
    'Row Count
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
     For i = 2 To RowCount
    
       'FInd tsv
       
         Total = Total + ws.Cells(i, 7).Value
       
       
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                  Ticker = ws.Cells(i, 1).Value
                
                
                     ws.Range("I" & Start).Value = Ticker
               
                     ws.Range("L" & Start).Value = Total
                

                Total = 0
                
                open_price = ws.Range("C" & PreviousStockPrice)
                
                close_price = ws.Range("F" & i)
                
                yearlychange = close_price - open_price
                
                ws.Range("J" & Start).Value = yearlychange
                
               
          
            'Percentage Change
                
            If open_price = 0 Then
                
                percentchange = 0
                
            Else
                open_price = ws.Range("C" & PreviousStockPrice)
              
                percentchange = yearlychange / open_price
                
            End If
            
            ws.Range("K" & Start).Value = percentchange
            
            ws.Range("K" & Start).NumberFormat = "0.00%"
            
                
            ' colors format change
                
               
            If ws.Range("J" & Start).Value >= 0 Then
                        ws.Range("j" & Start).Interior.ColorIndex = 4
            Else
                        ws.Range("j" & Start).Interior.ColorIndex = 3
                    
            End If
                'start of the next stock ticker
                Start = Start + 1
                
                PreviousStockPrice = i + 1
                
                
                
            End If
            
            Next i
            
             'loop for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume

            greatest_increase = 0
            greatest_decrease = 0
            gtv = 0
             
            'Set value of the last row for column K
            RowCount = ws.Cells(Rows.Count, "K").End(xlUp).Row
            
            For i = 2 To RowCount
            
            'First determine the Greatest Total Volume
            If ws.Range("L" & i).Value > gtv Then
               gtv = ws.Range("L" & i).Value
               ws.Range("R4").Value = gtv
               ws.Range("Q4").Value = ws.Range("I" & i).Value
               
            End If
            
            'Next determine Greatest % Increase
            If ws.Range("K" & i).Value > greatest_increase Then
                greatest_increase = ws.Range("K" & i).Value
                ws.Range("R2").Value = greatest_increase
                ws.Range("Q2").Value = ws.Range("I" & i).Value
                
            End If
            
            'Greatest % Decrease
            If ws.Range("K" & i).Value < greatest_decrease Then
                greatest_decrease = ws.Range("K" & i).Value
                ws.Range("R3").Value = greatest_decrease
                ws.Range("Q3").Value = ws.Range("I" & i).Value
                
            End If
            
            'Change format to "%"
            ws.Range("P2").NumberFormat = "0.00%"
            
            ws.Range("P3").NumberFormat = "0.00%"
    
        Next i


    Next ws

        
End Sub