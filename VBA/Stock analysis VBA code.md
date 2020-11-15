Sub stockticker()
Dim ws As Worksheet

'Loop Through All Sheets
For Each ws In Worksheets

    'Determine the variables
    Dim r As Long
    Dim LastRow As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim totalvolume As Double
    Dim y_change As Double
    Dim p_change As Double
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set the headers for summary table
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly change"
        ws.Range("K1") = "Percent change"
        ws.Range("L1") = "Total volume"
        
        reportCounter = 2  'set counter
        totalvolume = 0  'Initialize volume
        openprice = ws.Cells(2, 3).Value
        
        For r = 2 To LastRow
            
         If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            
            closeprice = ws.Cells(r, 6).Value
            y_change = closeprice - openprice
            totalvolume = totalvolume + ws.Cells(r, 7).Value  'calculate volume
            
            If openprice <> 0 Then
            p_change = (y_change / openprice) 'calculate yearly change and percent change
            
            End If
            
            'display summary values
            ws.Range("I" & reportCounter) = ws.Cells(r, 1).Value
            ws.Range("J" & reportCounter) = y_change
            ws.Range("K" & reportCounter) = p_change
            ws.Range("K" & reportCounter).NumberFormat = "0.00%"  'change number format in percent change column
            ws.Range("L" & reportCounter) = totalvolume
            
            If y_change > 0 Then
            ws.Range("J" & reportCounter).Interior.ColorIndex = 4
            Else
            ws.Range("J" & reportCounter).Interior.ColorIndex = 3
            End If
            
            
            'Reset variables for next ticker
            reportCounter = reportCounter + 1
            totalvolume = 0
            openprice = ws.Cells(r + 1, 3).Value
         Else
            totalvolume = totalvolume + ws.Cells(r, 7).Value
         End If
        Next r
    
    'Define variables for bonus summary
    Dim Summary_last_row As Long
    Dim i As Double
    Dim Maxincr As Double
    Dim Minincr As Double
    Dim Maxvol As Double
    Dim Best_tkr As String
    Dim Worst_tkr As String
    
    Maxincr = WorksheetFunction.Max(ws.Range("K:K"))
    Maxvol = WorksheetFunction.Max(ws.Range("L:L"))
    Minincr = WorksheetFunction.Min(ws.Range("K:K"))
    
    'Set headers
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    'Determine the last row for Percent change and total volume
    Summary_last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Display max and min values
    ws.Cells(2, 16).Value = Maxincr
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = Maxvol
    ws.Cells(3, 16).Value = Minincr
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
    For i = 2 To Summary_last_row
        Best_tkr = ws.Cells(i, 9).Value
        Worst_tkr = ws.Cells(i, 9).Value
        
        If ws.Cells(i, 11).Value = Maxincr Then
        ws.Cells(2, 15).Value = Best_tkr
    
        ElseIf ws.Cells(i, 11) = Minincr Then
        ws.Cells(3, 15).Value = Worst_tkr
        End If
        
        If ws.Cells(i, 12).Value = Maxvol Then
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    Next ws
End Sub
