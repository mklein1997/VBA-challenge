Sub homework()
    
    'Outline parameters and variables
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    Dim ticker As String
    
    Dim stockvolume As Double
    
    Dim summarytable As Long
    
    Dim openingprice As Double
    Dim closingprice As Double
    Dim previousprice As Double
    Dim annualchange As Double
    Dim percentchange As Double
    Dim lastRow As Long
    
    'Set initial values for variables
    stockvolume = 0
    summarytable = 2
    previousprice = 2
    percentchange = 0
    
    'Variable parameters for challenge portion
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As Double
    
    'Set values for challenge portion
    
    greatestincrease = 0
    greatestdecrease = 0
    greatestvolume = 0
    


    'Set headers for spreadsheet
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest_%_Increase"
    ws.Cells(3, 15).Value = "Greatest_%_Decrease"
    ws.Cells(4, 15).Value = "Greatest_Total_Volume"
    
    'Find last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through all stocks and pull unique values
    For i = 2 To lastRow
    
        'Add these values towards total stock volume for ticker
        stockvolume = stockvolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1).Value
        
        'pull ticker info and volume info for summary table
        ws.Range("I" & summarytable).Value = ticker
        
        ws.Range("L" & summarytable).Value = stockvolume
        
        'At the end of each ticker ID reset stock volume to 0
        
        stockvolume = 0
        
        'Set values of open close and annual change
        openingprice = ws.Range("C" & previousprice)
        closingprice = ws.Range("F" & i)
        
        annualchange = closingprice - openingprice
        
        ws.Range("J" & summarytable).Value = annualchange
        
         'Formula for percent change calculation
            If openingprice <> 0 Then
            openingprice = ws.Range("C" & previousprice)
            percentchange = (closingprice - openingprice) / openingprice
            
            End If
        
        'Show percent change in summary data and format as percentage
        
        ws.Range("K" & summarytable).Value = percentchange
        ws.Range("K" & summarytable).NumberFormat = "0.00%"
        
        'reset %change
        
        percentchange = 0
        summarytable = summarytable + 1
        previousprice = i + 1
        
        End If
        
        'Set cell color conditional formatting
        
        If ws.Range("J" & summarytable).Value >= 0 Then
            ws.Range("J" & summarytable).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summarytable).Interior.ColorIndex = 3
        End If
        
        
    Next i
    
    'challenge section
    'search for unique values
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set cell values and pull data to fill table
    
    For i = 2 To lastRow
    
    If ws.Range("K" & i).Value > ws.Cells(2, 17).Value Then
    ws.Cells(2, 17).Value = ws.Range("K" & i).Value
    ws.Cells(2, 16).Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("K" & i).Value < ws.Cells(3, 17).Value Then
    ws.Cells(3, 17).Value = ws.Range("K" & i).Value
    ws.Cells(3, 16).Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("L" & i).Value > ws.Cells(4, 17).Value Then
    ws.Cells(4, 17).Value = ws.Range("L" & i).Value
    ws.Cells(4, 16).Value = ws.Range("I" & i).Value
    
    End If
    
    Next i
    
    'Format for percentage
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
     
    
    
            
            
    Next ws
    

End Sub

    

