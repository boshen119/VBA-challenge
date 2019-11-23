Sub StockChanges()
    'Set variables
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim ws As Worksheet

        For Each ws In Worksheets
            'Set values for each Worksheet
            j = 0
            total = 0
            change = 0
            start = 2
            dailyChange = 0
    

            'Name titles for summary colums in each worksheet
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
                
                    'Get last row
                    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
                    
                    'Start the loop
                    For i = 2 To LastRow
                        
                'If ticker changes then print results
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Stores results in variables.
                    total = total + ws.Cells(i, 7).Value

                    'Handle zero total volume
                    If total = 0 Then
                        'Print the results
                        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = 0
                        ws.Range("K" & 2 + j).Value = "%" & 0
                        ws.Range("L" & 2 + j).Value = 0
                    Else
                        'Find First non zero starting value
                        If ws.Cells(start, 3) = 0 Then
                            For find_value = start To i
                                If ws.Cells(find_value, 3).Value <> 0 Then
                                    start = find_value
                                    Exit For
                                End If
                             Next find_value
                        End If
        
                        'Calculate Change
                        change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                        percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
                        
                        'Start of the next ticker symbol
                        start = i + 1
                        
                        'Print the results to separate worksheet
                        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = Round(change, 2)
                        ws.Range("K" & 2 + j).Value = "%" & percentChange
                        ws.Range("L" & 2 + j).Value = total
                        
                        'Colors positives green cells and negatives red cells
                        If change > 0 Then
                                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                            ElseIf change < 0 Then
                                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                            Else
                                ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                        End If
                        
                    End If
                
                    
                    'Reset variables for new stock ticker
                    total = 0
                    change = 0
                    j = j + 1
                    days = 0
                
                'If ticker is still the same add results
                    Else
                        total = total + ws.Cells(i, 7).Value
                End If
            
            Next i
            
            
            'Opening price subtract Closing price / opening price
            
            
            
             
    Next ws
    
    'Bonus: find Greatest % Increase across worksheets
    'Define each worksheet within workbook
    'Take values from Range I and Range K
    
    'i = 2
    'enter last row formula
    
    'Let's find the greatest % Increase!
    'Cell(2, 15).Value = "Greatest % Increase"
    'If Cells(i,i+1)> Cells(i,1) Then
    'Cells(
    
     'Cell(3, 15).Value = "Least % Increase"
End Sub
        
