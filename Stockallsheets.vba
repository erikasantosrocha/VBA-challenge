Sub Stockallsheets():

    For Each ws In Worksheets
    
    'Define variables
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim tickerCell As Long
    Dim lastrow As Long
    Dim lastRowUnique As Long
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
        
    'Get the WorksheetName
    WorksheetName = ws.Name
        
    'Define the range of the stock data from A1 to the last row of column A
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Define the location of the first ticker row
    tickerCell = 2
        
    'Assign to the first unique ticker the start row 2
    j = 2
        
    'Loop through each row in the stock data
        For i = 2 To lastrow
            
            'Identify if the ticker exists or is new
            If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                
                'Print the tickercell in column I if is new ticker
                ws.Cells(tickerCell, 9).value = ws.Cells(i, 1).value
                
                'Calculate the Yearly Change by substracting the last close amount - first open amount in column J
                ws.Cells(tickerCell, 10).value = ws.Cells(i, 6).value - ws.Cells(j, 3).value
                
                'Apply color to the yearly change cell, if is positive green and red for negative
                If ws.Cells(tickerCell, 10).value > 0 Then
                ws.Cells(tickerCell, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(tickerCell, 10).Interior.ColorIndex = 3
                End If
                    
                'Calculate the percent change (last close / first open) -1
                If ws.Cells(j, 3).value <> 0 Then
                percentChange = ((ws.Cells(i, 6).value / ws.Cells(j, 3).value) - 1)
                    
                'Print in column K the percent change
                ws.Cells(tickerCell, 11).value = percentChange
                Else
                ws.Cells(tickerCell, 11).value = 0
                End If

                'Calculate the total volume with sum function and print in column L
                ws.Cells(tickerCell, 12).value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Add to the previous ticker value the new loop value
                tickerCell = tickerCell + 1
                
                'Once is done with the unique ticker, assign the new start to the next ticker
                j = i + 1
                
            End If
            
        Next i
            
    'To show the summary table, will loop only through the Column I where the unique tickers information are located
            
    'Find last unique ticker cell in column I
    lastRowUnique = ws.Cells(Rows.Count, 9).End(xlUp).Row

        
    'Assign the first row (2) as default greates value before starts the loop
    greatestIncrease = ws.Cells(2, 11).value
    greatestDecrease = ws.Cells(2, 11).value
    greatestVolume = ws.Cells(2, 12).value

        
    'Loop through unique tickers to find the greatest increase, decrease and volume
        For i = 2 To lastRowUnique
                
            'Calculate greatest increase
            'If the next cell value is more than default value
            If ws.Cells(i, 11).value > greatestIncrease Then
            'select that cell value and print the ticker
            greatestIncrease = ws.Cells(i, 11).value
            ws.Cells(2, 16).value = ws.Cells(i, 9).value
                
            Else
            greatestIncrease = greatestIncrease
            End If
                
            'Calculate greatest decrease
            'If the next cell value is less than default value
            If ws.Cells(i, 11).value < greatestDecrease Then
            'select that cell value and print the ticker
            greatestDecrease = ws.Cells(i, 11).value
            ws.Cells(3, 16).value = ws.Cells(i, 9).value
                
            Else
            greatestDecrease = greatestDecrease
            End If
                
            'Calculate greatest volume
            'If the next cell value is more than default value
            If ws.Cells(i, 12).value > greatestVolume Then
            'select that cell value and print the ticker
            greatestVolume = ws.Cells(i, 12).value
            ws.Cells(4, 16).value = ws.Cells(i, 9).value
                
            Else
            greatestVolume = greatestVolume
            End If
                
            'Print the greatest increase, decrease and volume on column Q
            ws.Cells(2, 17).value = greatestIncrease
            ws.Cells(3, 17).value = greatestDecrease
            ws.Cells(4, 17).value = greatestVolume
            
        Next i
           
    'Add header to each column
        ws.Range("I1").value = "Ticker"
        ws.Range("J1").value = "Yearly Change"
        ws.Range("K1").value = "Percent Change"
        ws.Range("L1").value = "Total Stock Volume"
        ws.Range("P1").value = "Ticker"
        ws.Range("Q1").value = "Value"
        ws.Range("O2").value = "Greatest % Increase"
        ws.Range("O3").value = "Greatest % Decrease"
        ws.Range("O4").value = "Greatest Total Volume"
        
    'Add style to percent change cells
        ws.Columns("K").NumberFormat = "0.00%"
        
    ' Add style to greatest % cells
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    'Add style to all the columns for autofit
        ws.Columns("J:L").AutoFit
        ws.Columns("O:P").AutoFit
        
    Next ws
        
End Sub
