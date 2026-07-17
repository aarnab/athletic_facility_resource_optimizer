Sub GenerateEventBudgetDashboard()
    ' This macro creates a financial tracker for large-scale athletic trials
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
   
    ' PART 1: Generating the event budget data
   
    ' Creating Headers
    ws.Range("A1").Value = "Expense Category"
    ws.Range("B1").Value = "Allocated Budget (INR)"
    ws.Range("C1").Value = "Actual Cost (INR)"
    ws.Range("D1").Value = "Variance (Remaining)"
    
    ' Populating Rows with Event Data (e.g., Inter-NIT Trials)
    ws.Range("A2").Value = "Participant Refreshments"
    ws.Range("B2").Value = 15000
    ws.Range("C2").Value = 16500 ' Over budget
    
    ws.Range("A3").Value = "Lifting Accessories (Chalk, Belts)"
    ws.Range("B3").Value = 5000
    ws.Range("C3").Value = 4200
    
    ws.Range("A4").Value = "Tug of War Safety Mats"
    ws.Range("B4").Value = 8000
    ws.Range("C4").Value = 8000
    
    ws.Range("A5").Value = "Referee & Staff Honorariums"
    ws.Range("B5").Value = 12000
    ws.Range("C5").Value = 13500 ' Over budget


    ' PART 2: Automating the financial dashboard
   
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    
    ' Calculating the Variance (Allocated - Actual) for each row
    For i = 2 To lastRow
        ws.Cells(i, 4).Value = ws.Cells(i, 2).Value - ws.Cells(i, 3).Value
    Next i

    ' Applying bold headers and professional dark gray styling
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 28
    End With
    
    ' Highlighting deficits (where Variance is less than 0)
    Dim deficitTotal As Double
    deficitTotal = 0
    
    For i = 2 To lastRow
        If ws.Cells(i, 4).Value < 0 Then
            ' Painting the row light red to flag overspending
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 4)).Interior.Color = RGB(255, 204, 204)
            ' Adding to the total deficit tracker
            deficitTotal = deficitTotal + Abs(ws.Cells(i, 4).Value)
        ElseIf ws.Cells(i, 4).Value > 0 Then
            ' Painting the row light green to show a healthy surplus
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 4)).Interior.Color = RGB(204, 255, 204)
        End If
    Next i
    
    ' Applying currency formatting to the financial columns
    ws.Range("B2:D" & lastRow).NumberFormat = "?#,##0"
    
    ' Output final financial analysis alert
    If deficitTotal > 0 Then
        MsgBox "Warning: The event is currently over budget by ?" & deficitTotal & ". Please review the flagged red categories.", vbExclamation, "Budget Alert"
    Else
        MsgBox "Event budget is optimized and operating within limits.", vbInformation, "Budget Status"
    End If
    
End Sub
