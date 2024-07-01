Sub stock_organizer()
    'Setting dimensions
    Dim total As Double
    Dim rowindex As Long
    Dim change As Double
    Dim columnindex As Integer
    Dim start As Long
    Dim rowcount As Long
    Dim percentchange As Double
    Dim days As Integer
    Dim dailychange As Single
    Dim averagechange As Double
    Dim ws As Worksheet
    Dim find_value As Double

    'Resetting the values
    For Each ws In Worksheets
        columnindex = 0
        total = 0
        change = 0
        start = 2
        dailychange = 0
        
        'Setting up titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"

        'Obtaining the number of last row
        rowcount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For rowindex = 2 To rowcount
            'If the ticker changes then we print the results
            If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
                'Storing our results in the variable
                total = total + ws.Cells(rowindex, 7).Value
                
                If total = 0 Then
                    'Placing results
                    ws.Range("I" & 2 + columnindex).Value = ws.Cells(rowindex, 1).Value
                    ws.Range("J" & 2 + columnindex).Value = 0
                    ws.Range("K" & 2 + columnindex).Value = "%" & 0
                    ws.Range("L" & 2 + columnindex).Value = 0
                Else
                    If ws.Cells(start, 3).Value = 0 Then
                        For find_value = start To rowindex
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    change = ws.Cells(rowindex, 6).Value - ws.Cells(start, 3).Value
                    
                    If ws.Cells(start, 3).Value <> 0 Then
                        percentchange = change / ws.Cells(start, 3).Value
                    Else
                        percentchange = 0
                    End If
                    
                    start = rowindex + 1
                    
                    ws.Range("I" & 2 + columnindex).Value = ws.Cells(rowindex, 1).Value
                    ws.Range("J" & 2 + columnindex).Value = change
                    ws.Range("J" & 2 + columnindex).NumberFormat = "0.00"
                    ws.Range("K" & 2 + columnindex).Value = percentchange
                    ws.Range("K" & 2 + columnindex).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + columnindex).Value = total
                    
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 0
                    End Select
                End If
                
                total = 0
                change = 0
                columnindex = columnindex + 1
                days = 0
                dailychange = 0
            Else
                total = total + ws.Cells(rowindex, 7).Value
            End If
        Next rowindex

        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))

        Dim increase_number As Long, decrease_number As Long, volume_number As Long
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)

        ws.Range("P2").Value = ws.Cells(increase_number + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(decrease_number + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(volume_number + 1, 9).Value
    Next ws
End Sub
