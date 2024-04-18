Attribute VB_Name = "Module1"
Sub stockAnalysis()

    Dim total As Double ' Total stock volume
    Dim row As Long      ' Loop control
    Dim rowCount As Long    ' Number of rows
    Dim yearlyChange As Double   '  yearly change for ticker
    Dim percentChange As Double ' Percent change
    Dim summaryTableRow As Long ' Row of the summary table
    Dim stockStartRow As Long       ' Holds start of stock
    Dim openPrice As Double ' Opening price for ticker
    
    'For each ws(hopefully)
    For Each ws In Worksheets
    
        ' header rows
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        summaryTableRow = 0
        total = 0
        yearlyChange = 0
        stockStartRow = 2
        openPrice = 0 ' openPrice variable
    
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
        For row = 2 To rowCount
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                total = total + ws.Cells(row, 7).Value
    
                If total = 0 Then
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                Else
                    If ws.Cells(stockStartRow, 3).Value = 0 Then
                        For findValue = stockStartRow To row
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                stockStartRow = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                
                    yearlyChange = (ws.Cells(row, 6).Value - openPrice) ' Calc yearly change
                    If openPrice <> 0 Then
                        percentChange = yearlyChange / openPrice ' Calc percentage change
                    Else
                        percentChange = 0
                    End If
                    
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
                    ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###"
                
                    If yearlyChange > 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                    ElseIf yearlyChange < 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    End If
                    
                End If
    
                total = 0 ' Reset total for the next stock
                yearlyChange = 0 ' Reset yearly change for the next stock
                summaryTableRow = summaryTableRow + 1
                openPrice = 0 ' Reset openPrice for the next stock
            
            Else
                total = total + ws.Cells(row, 7).Value
                
                If openPrice = 0 Then ' Update openPrice if it's zero
                    openPrice = ws.Cells(row, 3).Value
                End If
            
            End If
    
        Next row
        
    Next ws
    
    
    For Each ws In Worksheets
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
        ws.Range("Q4").NumberFormat = "#,###"
        
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
        
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
        
        volNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        ws.Range("P4").Value = ws.Cells(volNumber + 1, 9)
    Next ws
    
End Sub
