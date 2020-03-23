Attribute VB_Name = "Module1"
Sub vba_challenge():

For Each ws In Worksheets

    'stock name
    Dim ticker As String
    
    'opening price at beginning of the year
    Dim open_price As Double
    
    'closing price at the end of the year
    Dim close_price As Double
    
    'total volume of stocks traded all year
    Dim volume As LongLong
    
    'row counter to keep track of where we're writing the output
    Dim counter As Integer
    counter = 2
               
    'variable to hold the percent change
    Dim percent_change As Double
    
    Dim greates_increase_ticker As String
    Dim greatest_increase As Double
    greatest_increase = 0
    Dim greatest_decrease_ticker As String
    Dim greatest_decrease As Double
    greatest_increase = 0
    Dim greatest_volume_ticker As String
    Dim greatest_volume As Long
    greatest_volume = 0
    
    
    'initializing the opening price for the first ticker name
    open_price = ws.Cells(2, 3).Value
        
    'writing column names for the output
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 10).Value = "yearly change"
    ws.Cells(1, 11).Value = "percent change"
    ws.Cells(1, 12).Value = "total stock volume"
    
    'iterate through all the rows
    For i = 2 To ws.Rows.Count
        
        'check for values
        If IsEmpty(ws.Cells(i, 1)) = True Then
            Exit For
        End If
                
        'check if under same ticker name
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            'set and write ticker value
            ticker = ws.Cells(i, 1).Value
            ws.Cells(counter, 9).Value = ticker
            
            'set final close price and do calculations with open price
            close_price = ws.Cells(i, 6).Value
            
            'yearly change
            ws.Cells(counter, 10).Value = close_price - open_price
            
            'percent change
            If open_price = 0 Then
                percent_change = close_price - open_price
            Else
                percent_change = (close_price - open_price) / open_price
            End If
            
            ws.Cells(counter, 11).Value = percent_change
            
             If ws.Cells(counter, 11).Value > 0 Then
                ws.Cells(counter, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(counter, 11).Value < 0 Then
                ws.Cells(counter, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(counter, 11).Interior.ColorIndex = 2
            End If
            
            ws.Cells(counter, 11).NumberFormat = "0.00%"
           
            
            'add last volume count and write volume
            volume = ws.Cells(i, 7).Value
            ws.Cells(counter, 12).Value = volume
            
            'update counter
            counter = counter + 1
            
            'set open price for next ticker
            open_price = ws.Cells((i + 1), 3).Value
        
        Else
            'sum volume
            volume = volume + ws.Cells(i, 7).Value
                        
        End If
          
    Next i
    
     'iterate through all the rows
    For i = 2 To ws.Rows.Count
        
        'check for values
        If IsEmpty(ws.Cells(i, 9)) = True Then
            Exit For
        End If
        
        'compare percent change
        If ws.Cells(i, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(i, 11).Value
            greatest_increase_ticker = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11) < greatest_decrease Then
            greatest_decrease = ws.Cells(i, 11).Value
            greatest_decrease_ticker = ws.Cells(i, 9).Value
        End If
        
        'compare volume
        If ws.Cells(i, 12).Value > greatest_volume Then
            greatest_volume = ws.Cells(i, 12).Value
            greatest_volume_ticker = ws.Cells(i, 9).Value
        End If
            
     Next i
    
    'write all the garbage from the previous for loop
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
    
     ws.Cells(2, 15).Value = "Greatest % increase"
     ws.Cells(2, 16).Value = greatest_increase_ticker
     ws.Cells(2, 17).Value = greatest_increase
     ws.Cells(2, 17).NumberFormat = "0.00%"
     
     ws.Cells(3, 15).Value = "Greatest % decrease"
     ws.Cells(3, 16).Value = greatest_decrease_ticker
     ws.Cells(3, 17).Value = greatest_decrease
     ws.Cells(3, 17).NumberFormat = "0.00%"
     
     ws.Cells(4, 15).Value = "Greatest Total Volume"
     ws.Cells(4, 16).Value = greatest_volume_ticker
     ws.Cells(4, 17).Value = greatest_volume
     
    
    
Next ws
           
        
End Sub

