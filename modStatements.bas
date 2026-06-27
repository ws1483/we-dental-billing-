Attribute VB_Name = "modStatements"
Option Explicit

' =====================================================
' ===== STATEMENT LAYOUT CONSTANTS =====
' =====================================================

Private Const STMT_TITLE_ROW As Long = 6
Private Const STMT_HEADER_START As Long = 8
Private Const STMT_TABLE_HEADER As Long = 14
Private Const STMT_FIRST_DATA As Long = 15

Private Const STMT_FOOTER_ROW_COUNT As Long = 9
Private Const STMT_AGING_BLOCK_ROWS As Long = 3

Private Type AgingBuckets
    current As Double
    days30 As Double
    days60 As Double
    days90 As Double
End Type

' =====================================================
' ===== GENERATE STATEMENT =====
' =====================================================

Public Sub GenerateStatementWithParams(ByVal customerName As String, ByVal fromDate As Date, ByVal toDate As Date)
    Dim wsStmt As Worksheet, wbTrack As Workbook, wsTrack As Worksheet
    Dim trackingPath As String
    Dim lastTrackRow As Long, r As Long, outRow As Long
    
    Dim invNum As String, invDate As Variant, cust As String
    Dim patient As String, appType As String
    Dim debitAmt As Double, creditAmt As Double, paidAmt As Double, lineBalance As Double
    Dim runningBalance As Double, status As String, dInvDate As Date
    
    Dim totalDebits As Double, totalPaid As Double, totalCredits As Double
    Dim invoiceCount As Long
    
    Dim aging As AgingBuckets
    
    If Not SheetExists(SHEET_STATEMENT) Then
        ShowError "Statement sheet not found."
        Exit Sub
    End If
    
    Set wsStmt = ThisWorkbook.Sheets(SHEET_STATEMENT)
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Generating statement..."
    
    ClearStatementBody wsStmt
    PopulateCustomerHeader wsStmt, customerName, fromDate, toDate
    WriteStatementTableHeaders wsStmt
    SetStatementColumnWidths wsStmt
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then GoTo CleanUp
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then GoTo CleanUp
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        ShowWarning "No invoice tracking data found.", "No Data"
        GoTo CleanUp
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    lastTrackRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row
    
    outRow = STMT_FIRST_DATA
    totalDebits = 0
    totalPaid = 0
    totalCredits = 0
    runningBalance = 0
    invoiceCount = 0
    
    aging.current = 0
    aging.days30 = 0
    aging.days60 = 0
    aging.days90 = 0
    
    For r = 2 To lastTrackRow
        invNum = Trim(CStr(wsTrack.Cells(r, COL_I_NUMBER).Value))
        invDate = wsTrack.Cells(r, COL_I_DATE).Value
        cust = Trim(CStr(wsTrack.Cells(r, COL_I_CUSTOMER).Value))
        
        If cust <> "" Then
            If StrComp(cust, customerName, vbTextCompare) = 0 Then
                If IsDate(invDate) Then
                    dInvDate = CDate(invDate)
                    
                    If dInvDate >= fromDate And dInvDate <= toDate Then
                        patient = Trim(CStr(wsTrack.Cells(r, COL_I_PATIENT).Value))
                        appType = Trim(CStr(wsTrack.Cells(r, COL_I_APPLIANCE).Value))
                        debitAmt = r2(wsTrack.Cells(r, COL_I_TOTAL).Value)
                        creditAmt = r2(wsTrack.Cells(r, COL_I_CREDITAMT).Value)
                        status = Trim(CStr(wsTrack.Cells(r, COL_I_STATUS).Value))
                        
                        paidAmt = r2(wsTrack.Cells(r, COL_I_PAIDAMOUNT).Value)
                        
                        If paidAmt = 0 And LCase$(status) = LCase$(STATUS_INVOICE_PAID) Then
                            paidAmt = Round(debitAmt - creditAmt, 2)
                            If paidAmt < 0 Then paidAmt = 0
                        End If
                        
                        lineBalance = Round(debitAmt - paidAmt - creditAmt, 2)
                        If lineBalance < 0 Then lineBalance = 0
                        
                        runningBalance = Round(runningBalance + lineBalance, 2)
                        
                        WriteStatementLine wsStmt, outRow, dInvDate, invNum, patient, appType, _
                                           debitAmt, paidAmt, creditAmt, lineBalance, status
                        
                        totalDebits = Round(totalDebits + debitAmt, 2)
                        totalPaid = Round(totalPaid + paidAmt, 2)
                        totalCredits = Round(totalCredits + creditAmt, 2)
                        invoiceCount = invoiceCount + 1
                        
                        If lineBalance > 0 Then
                            AddToAgingBucket aging, dInvDate, lineBalance
                        End If
                        
                        outRow = outRow + 1
                    End If
                End If
            End If
        End If
    Next r
    
    CloseTrackingWorkbook wbTrack, False
    
    If invoiceCount = 0 Then
        wsStmt.Cells(STMT_FIRST_DATA, "A").Value = "No invoices found for this period."
        wsStmt.Range("A" & STMT_FIRST_DATA & ":H" & STMT_FIRST_DATA).Merge
        wsStmt.Cells(STMT_FIRST_DATA, "A").HorizontalAlignment = xlCenter
        outRow = STMT_FIRST_DATA + 1
    End If
    
    Dim totalsRow As Long
    totalsRow = outRow + 1
    
    Dim balanceDue As Double, totalExcl As Double, totalVat As Double
    balanceDue = Round(totalDebits - totalPaid - totalCredits, 2)
    If balanceDue < 0 Then balanceDue = 0
    
    If balanceDue > 0 Then
        totalExcl = Round(balanceDue / (1 + VAT_RATE), 2)
        totalVat = Round(balanceDue - totalExcl, 2)
    Else
        totalExcl = 0
        totalVat = 0
    End If
    
    DrawTotalsSeparatorLine wsStmt, totalsRow
    WriteStatementTotals wsStmt, totalsRow, totalDebits, totalPaid, totalCredits, totalExcl, totalVat, balanceDue
    
    Dim agingRow As Long
    agingRow = (totalsRow + 5) + 2
    WriteAgingSummary wsStmt, agingRow, aging
    
    Dim footerStartRow As Long
    footerStartRow = agingRow + STMT_AGING_BLOCK_ROWS
    WriteStatementFooter wsStmt, footerStartRow
    
    Dim lastPrintRow As Long
    lastPrintRow = footerStartRow + STMT_FOOTER_ROW_COUNT + 1
    
    SetupStatementPrint wsStmt, lastPrintRow
    
    wsStmt.Activate
    wsStmt.Range("A1").Select
    
    RestoreExcelState
    
    ShowInfo "Statement generated for " & customerName & vbNewLine & _
             "Period: " & Format(fromDate, "YYYY/MM/DD") & " to " & Format(toDate, "YYYY/MM/DD") & vbNewLine & _
             "Invoices found: " & invoiceCount & vbNewLine & _
             "Total Invoiced: " & FmtR(totalDebits) & vbNewLine & _
             "Total Paid: " & FmtR(totalPaid) & vbNewLine & _
             "Total Credits: " & FmtR(totalCredits) & vbNewLine & _
             "Balance Due: " & FmtR(balanceDue) & vbNewLine & vbNewLine & _
             "Aging:" & vbNewLine & _
             "  Current : " & FmtR(aging.current) & vbNewLine & _
             "  30 Days : " & FmtR(aging.days30) & vbNewLine & _
             "  60 Days : " & FmtR(aging.days60) & vbNewLine & _
             "  90+ Days: " & FmtR(aging.days90), _
             "Statement Complete"
    Exit Sub

CleanUp:
    RestoreExcelState
    Exit Sub

ErrorHandler:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source
    
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    RestoreExcelState
    On Error GoTo 0
    
    ShowError "Error generating statement." & vbNewLine & vbNewLine & _
              "Error #" & errNum & " (" & errSrc & "):" & vbNewLine & _
              IIf(Len(errDesc) > 0, errDesc, "(no description provided by Excel)")
End Sub

' =====================================================
' ===== AGING BUCKET =====
' =====================================================

Private Sub AddToAgingBucket(ByRef buckets As AgingBuckets, ByVal invDate As Date, ByVal amount As Double)
    Dim ageDays As Long
    ageDays = DateDiff("d", invDate, Date)
    
    If ageDays < 0 Then ageDays = 0
    
    Select Case ageDays
        Case 0 To 30
            buckets.current = Round(buckets.current + amount, 2)
        Case 31 To 60
            buckets.days30 = Round(buckets.days30 + amount, 2)
        Case 61 To 90
            buckets.days60 = Round(buckets.days60 + amount, 2)
        Case Else
            buckets.days90 = Round(buckets.days90 + amount, 2)
    End Select
End Sub

Private Sub WriteAgingSummary(ws As Worksheet, ByVal headerRow As Long, ByRef buckets As AgingBuckets)
    Dim amountRow As Long
    amountRow = headerRow + 1
    
    Dim cGreen As Long, cAmber As Long, cOrange As Long, cred As Long
    cGreen = RGB(0, 128, 0)
    cAmber = RGB(204, 153, 0)
    cOrange = RGB(230, 100, 0)
    cred = RGB(192, 0, 0)
    
    On Error Resume Next
    
    With ws
        .Range("A" & headerRow & ":D" & headerRow).UnMerge
        .Range("A" & headerRow & ":D" & headerRow).Merge
        .Cells(headerRow, "A").Value = "Aging Summary"
        .Cells(headerRow, "A").Font.bold = True
        .Cells(headerRow, "A").Font.Size = 10
        .Cells(headerRow, "A").HorizontalAlignment = xlLeft
        .Cells(headerRow, "A").VerticalAlignment = xlCenter
        
        .Cells(headerRow, "E").Value = "Current"
        .Cells(headerRow, "F").Value = "30 Days"
        .Cells(headerRow, "G").Value = "60 Days"
        .Cells(headerRow, "H").Value = "90+ Days"
        
        With .Range("E" & headerRow & ":H" & headerRow)
            .Font.bold = True
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        
        .Cells(headerRow, "E").Font.Color = cGreen
        .Cells(headerRow, "F").Font.Color = cAmber
        .Cells(headerRow, "G").Font.Color = cOrange
        .Cells(headerRow, "H").Font.Color = cred
        
        .Cells(amountRow, "E").Value = Round(buckets.current, 2)
        .Cells(amountRow, "F").Value = Round(buckets.days30, 2)
        .Cells(amountRow, "G").Value = Round(buckets.days60, 2)
        .Cells(amountRow, "H").Value = Round(buckets.days90, 2)
        
        With .Range("E" & amountRow & ":H" & amountRow)
            .NumberFormat = """R""#,##0.00"
            .Font.bold = True
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
        
        .Cells(amountRow, "E").Font.Color = cGreen
        .Cells(amountRow, "F").Font.Color = cAmber
        .Cells(amountRow, "G").Font.Color = cOrange
        .Cells(amountRow, "H").Font.Color = cred
        
        With .Range("E" & headerRow & ":H" & amountRow)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
        End With
    End With
    
    On Error GoTo 0
End Sub

Private Sub SetStatementColumnWidths(ws As Worksheet)
    On Error Resume Next
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 32
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 13
    On Error GoTo 0
End Sub

Private Sub ClearStatementBody(ws As Worksheet)
    Dim lastRow As Long
    
    On Error Resume Next
    
    ws.Range("A8:H13").UnMerge
    ws.Range("A13:H13").ClearContents
    
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastRow < STMT_FIRST_DATA Then lastRow = STMT_FIRST_DATA + 50
    
    If lastRow >= STMT_TABLE_HEADER Then
        With ws.Range("A" & STMT_TABLE_HEADER & ":H" & lastRow + 30)
            .ClearContents
            .Interior.Pattern = xlNone
            .Borders.LineStyle = xlNone
            .Font.bold = False
            .Font.Color = vbBlack
            .Font.Size = 10
            .UnMerge
        End With
    End If
    
    On Error GoTo 0
End Sub

' =====================================================
' ===== CUSTOMER HEADER LOOKUP (LEFT-ALIGNED LABELS) =====
' =====================================================

Private Sub PopulateCustomerHeader(ws As Worksheet, _
                                   ByVal customerName As String, _
                                   ByVal fromDate As Date, _
                                   ByVal toDate As Date)
    Dim wsCust As Worksheet
    Dim custRow As Long
    Dim ambiguous As Boolean
    
    On Error Resume Next     ' header writes must NEVER abort the statement
    
    ' --- Step 1: wipe header fields so prior data can't leak ---
    ws.Range("A8:H13").UnMerge
    ws.Range("A8:H13").ClearContents
    
    ' --- Step 2: re-merge value cells so long values have room ---
    ws.Range("B8:D8").Merge
    ws.Range("B9:D9").Merge
    ws.Range("B10:D10").Merge
    ws.Range("B12:D12").Merge
    ws.Range("F8:H8").Merge
    ws.Range("F9:H9").Merge
    ws.Range("F10:H10").Merge
    ws.Range("F11:H11").Merge
    ws.Range("E12:H12").Merge
    ws.Range("B13:D13").Merge
    
    ' --- Step 3: write left-hand labels (column A) ---
    ws.Range("A8").Value = "Bill To:"
    ws.Range("A9").Value = "Practice:"
    ws.Range("A10").Value = "Adress:"
    ws.Range("A12").Value = "PostCode:"
    ws.Range("A13").Value = "VAT No:"
    
    ' --- Step 4: write right-hand labels (column E) ---
    ws.Range("E8").Value = "Date :"
    ws.Range("E9").Value = "Period:"
    ws.Range("E10").Value = "Cust ID:"
    ws.Range("E11").Value = "Practice No:"
    ws.Range("D12").Value = "Email:"
    
    ' Always-present values
    ws.Range("B8").Value = customerName
    ws.Range("F8").Value = Format(Date, "YYYY/MM/DD")
    ws.Range("F9").Value = Format(fromDate, "YYYY/MM/DD") & " - " & Format(toDate, "YYYY/MM/DD")
    
    ' --- Step 5: formatting ---
    With ws.Range("A8:A13")
        .Font.bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("E8:E11")
        .Font.bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("D12")
        .Font.bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("B8:D8,B9:D9,B10:D10,B12:D12,B13:D13,F8:H8,F9:H9,F10:H10,F11:H11,E12:H12")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    With ws.Range("B8")
        .Font.bold = True
        .Font.Size = 11
    End With
    
    ' --- Step 6: Look up customer details ---
    On Error GoTo DoneHeader
    
    If Not SheetExists(SHEET_CUSTOMERS) Then GoTo DoneHeader
    Set wsCust = ThisWorkbook.Sheets(SHEET_CUSTOMERS)
    
    custRow = FindCustomerRow(customerName, ambiguous)
    
    If ambiguous Then
        ShowWarning "Multiple customer rows match '" & customerName & "' on the Customers sheet." & vbNewLine & _
                    "The statement header may be incomplete. Please confirm the customer name is unique.", _
                    "Ambiguous Customer Match"
    End If
    
    If custRow = 0 Then GoTo DoneHeader
    
    ws.Range("B9").Value = SafeString(wsCust.Cells(custRow, "B").Value)
    ws.Range("B10").Value = SafeString(wsCust.Cells(custRow, "C").Value)
    ws.Range("B12").Value = SafeString(wsCust.Cells(custRow, "D").Value)
    ws.Range("F10").Value = SafeString(wsCust.Cells(custRow, "J").Value)
    ws.Range("F11").Value = SafeString(wsCust.Cells(custRow, "G").Value)
    ws.Range("E12").Value = SafeString(wsCust.Cells(custRow, "I").Value)
    ws.Range("B13").Value = SafeString(wsCust.Cells(custRow, "K").Value)
    
DoneHeader:
    On Error GoTo 0
End Sub

' Now returns ambiguity info so the caller can warn the user.
' Prefers exact match. Only falls back to InStr-match when there is
' exactly ONE such candidate (otherwise it returns 0 + ambiguous=True).
Private Function FindCustomerRow(ByVal customerName As String, _
                                  ByRef ambiguous As Boolean) As Long
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim target As String, candidate As String
    Dim partialHits As Long, partialRow As Long
    
    FindCustomerRow = 0
    ambiguous = False
    
    If Not SheetExists(SHEET_CUSTOMERS) Then Exit Function
    Set ws = ThisWorkbook.Sheets(SHEET_CUSTOMERS)
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Function
    
    target = NormaliseName(customerName)
    If target = "" Then Exit Function
    
    ' --- Pass 1: exact match wins immediately ---
    For i = 2 To lastRow
        candidate = NormaliseName(CStr(ws.Cells(i, "A").Value))
        If candidate <> "" Then
            If StrComp(candidate, target, vbTextCompare) = 0 Then
                FindCustomerRow = i
                Exit Function
            End If
        End If
    Next i
    
    ' --- Pass 2: fuzzy fallback, but only if UNIQUE ---
    partialHits = 0
    partialRow = 0
    
    For i = 2 To lastRow
        candidate = NormaliseName(CStr(ws.Cells(i, "A").Value))
        If candidate <> "" Then
            If InStr(1, candidate, target, vbTextCompare) > 0 _
               Or InStr(1, target, candidate, vbTextCompare) > 0 Then
                partialHits = partialHits + 1
                partialRow = i
                If partialHits > 1 Then Exit For
            End If
        End If
    Next i
    
    If partialHits = 1 Then
        FindCustomerRow = partialRow
    ElseIf partialHits > 1 Then
        ambiguous = True
    End If
End Function

Private Function NormaliseName(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    
    t = Replace(t, Chr(160), " ")
    t = Replace(t, ".", "")
    
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    
    NormaliseName = Trim(t)
End Function

Private Sub WriteStatementTableHeaders(ws As Worksheet)
    Dim headerRow As Long
    headerRow = STMT_TABLE_HEADER
    
    With ws
        .Cells(headerRow, "A").Value = "Date"
        .Cells(headerRow, "B").Value = "Invoice #"
        .Cells(headerRow, "C").Value = "Patient"
        .Cells(headerRow, "D").Value = "Type"
        .Cells(headerRow, "E").Value = "Invoiced"
        .Cells(headerRow, "F").Value = "Paid"
        .Cells(headerRow, "G").Value = "Credit"
        .Cells(headerRow, "H").Value = "Balance"
        
        With .Range("A" & headerRow & ":H" & headerRow)
            .Font.bold = True
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
        
        .Cells(headerRow, "A").HorizontalAlignment = xlLeft
        .Cells(headerRow, "B").HorizontalAlignment = xlLeft
        .Cells(headerRow, "C").HorizontalAlignment = xlLeft
        .Cells(headerRow, "D").HorizontalAlignment = xlLeft
        .Cells(headerRow, "E").HorizontalAlignment = xlRight
        .Cells(headerRow, "F").HorizontalAlignment = xlRight
        .Cells(headerRow, "G").HorizontalAlignment = xlRight
        .Cells(headerRow, "H").HorizontalAlignment = xlRight
    End With
End Sub

Private Sub WriteStatementLine(ws As Worksheet, ByVal outRow As Long, _
                               ByVal invDate As Date, _
                               ByVal invNum As String, ByVal patient As String, _
                               ByVal appType As String, ByVal debitAmt As Double, _
                               ByVal paidAmt As Double, ByVal creditAmt As Double, _
                               ByVal lineBalance As Double, ByVal status As String)
    With ws
        .Cells(outRow, "A").Value = invDate
        .Cells(outRow, "A").NumberFormat = "YYYY/MM/DD"
        .Cells(outRow, "B").Value = invNum
        .Cells(outRow, "C").Value = patient
        .Cells(outRow, "D").Value = appType
        
        .Cells(outRow, "E").Value = Round(debitAmt, 2)
        .Cells(outRow, "E").NumberFormat = """R""#,##0.00"
        
        If paidAmt > 0 Then
            .Cells(outRow, "F").Value = Round(paidAmt, 2)
            .Cells(outRow, "F").NumberFormat = """R""#,##0.00"
            .Cells(outRow, "F").Font.Color = RGB(0, 128, 0)
        Else
            .Cells(outRow, "F").Value = ""
        End If
        
        If creditAmt > 0 Then
            .Cells(outRow, "G").Value = Round(creditAmt, 2)
            .Cells(outRow, "G").NumberFormat = """R""#,##0.00"
            .Cells(outRow, "G").Font.Color = RGB(0, 128, 0)
        Else
            .Cells(outRow, "G").Value = ""
        End If
        
        .Cells(outRow, "H").Value = Round(lineBalance, 2)
        .Cells(outRow, "H").NumberFormat = """R""#,##0.00"
        
        If LCase$(status) = LCase$(STATUS_INVOICE_PAID) Then
            .Range("A" & outRow & ":H" & outRow).Font.Color = RGB(128, 128, 128)
            .Cells(outRow, "H").Value = 0
            .Cells(outRow, "H").NumberFormat = """R""#,##0.00"
            .Cells(outRow, "H").Font.Color = RGB(0, 128, 0)
        ElseIf LCase$(status) = LCase$(STATUS_INVOICE_PARTIAL) Then
            .Cells(outRow, "H").Font.Color = RGB(204, 102, 0)
            .Cells(outRow, "H").Font.bold = True
        End If
        
        If lineBalance > 0 And LCase$(status) <> LCase$(STATUS_INVOICE_PARTIAL) Then
            .Cells(outRow, "H").Font.bold = True
        End If
        
        .Range("A" & outRow & ":H" & outRow).Borders.LineStyle = xlNone
        .Range("A" & outRow & ":H" & outRow).VerticalAlignment = xlCenter
        
        .Range("A" & outRow & ":D" & outRow).HorizontalAlignment = xlLeft
        .Range("E" & outRow & ":H" & outRow).HorizontalAlignment = xlRight
    End With
End Sub

Private Sub DrawTotalsSeparatorLine(ws As Worksheet, ByVal totalsRow As Long)
    Dim lineRow As Long
    lineRow = totalsRow - 1
    If lineRow < STMT_TABLE_HEADER Then lineRow = totalsRow
    
    On Error Resume Next
    With ws.Range("A" & lineRow & ":H" & lineRow).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    On Error GoTo 0
End Sub

' Always renders all 6 totals rows so layout is consistent regardless
' of whether credits are present. Fixes the "blank row gap" bug.
Private Sub WriteStatementTotals( _
    ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal totalDebits As Double, _
    ByVal totalPaid As Double, _
    ByVal totalCredits As Double, _
    ByVal totalExcl As Double, _
    ByVal vatAmount As Double, _
    ByVal balanceDue As Double)
    
    Dim rowIdx As Long
    
    With ws
        ' Row 0: Total Invoiced
        rowIdx = startRow
        .Cells(rowIdx, "F").Value = "Total Invoiced"
        .Cells(rowIdx, "H").Value = Round(totalDebits, 2)
        .Cells(rowIdx, "H").NumberFormat = """R""#,##0.00"
        .Range("F" & rowIdx & ":H" & rowIdx).Font.bold = True
        .Range("F" & rowIdx & ":H" & rowIdx).HorizontalAlignment = xlRight
        
        ' Row 1: Total Paid
        rowIdx = startRow + 1
        .Cells(rowIdx, "F").Value = "Total Paid"
        .Cells(rowIdx, "H").Value = Round(totalPaid, 2)
        .Cells(rowIdx, "H").NumberFormat = """R""#,##0.00"
        .Range("F" & rowIdx & ":H" & rowIdx).Font.bold = True
        .Range("F" & rowIdx & ":H" & rowIdx).Font.Color = RGB(0, 128, 0)
        .Range("F" & rowIdx & ":H" & rowIdx).HorizontalAlignment = xlRight
        
        ' Row 2: Total Credits — ALWAYS rendered so subsequent rows
        ' stay at fixed offsets (no blank gap when credits are zero).
        rowIdx = startRow + 2
        .Cells(rowIdx, "F").Value = "Total Credits"
        .Cells(rowIdx, "H").Value = Round(totalCredits, 2)
        .Cells(rowIdx, "H").NumberFormat = """R""#,##0.00"
        .Range("F" & rowIdx & ":H" & rowIdx).Font.bold = True
        If totalCredits > 0 Then
            .Range("F" & rowIdx & ":H" & rowIdx).Font.Color = RGB(0, 128, 0)
        Else
            .Range("F" & rowIdx & ":H" & rowIdx).Font.Color = RGB(128, 128, 128)
        End If
        .Range("F" & rowIdx & ":H" & rowIdx).HorizontalAlignment = xlRight
        
        ' Row 3: separator + Outstanding Excl
        rowIdx = startRow + 3
        .Range("F" & rowIdx & ":H" & rowIdx).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("F" & rowIdx & ":H" & rowIdx).Borders(xlEdgeTop).Weight = xlThin
        .Cells(rowIdx, "F").Value = "Outstanding Excl"
        .Cells(rowIdx, "H").Value = Round(totalExcl, 2)
        .Cells(rowIdx, "H").NumberFormat = """R""#,##0.00"
        .Range("F" & rowIdx & ":H" & rowIdx).Font.bold = True
        .Range("F" & rowIdx & ":H" & rowIdx).HorizontalAlignment = xlRight
        
        ' Row 4: VAT
        rowIdx = startRow + 4
        .Cells(rowIdx, "F").Value = "VAT"
        .Cells(rowIdx, "H").Value = Round(vatAmount, 2)
        .Cells(rowIdx, "H").NumberFormat = """R""#,##0.00"
        .Range("F" & rowIdx & ":H" & rowIdx).Font.bold = True
        .Range("F" & rowIdx & ":H" & rowIdx).HorizontalAlignment = xlRight
        
        ' Row 5: BALANCE DUE
        rowIdx = startRow + 5
        .Cells(rowIdx, "F").Value = "BALANCE DUE"
        .Cells(rowIdx, "H").Value = Round(balanceDue, 2)
        .Cells(rowIdx, "H").NumberFormat = """R""#,##0.00"
        .Range("F" & rowIdx & ":H" & rowIdx).Font.bold = True
        .Range("F" & rowIdx & ":H" & rowIdx).Font.Size = 11
        .Range("F" & rowIdx & ":H" & rowIdx).HorizontalAlignment = xlRight
        
        If balanceDue > 0 Then
            .Range("F" & rowIdx & ":H" & rowIdx).Font.Color = RGB(192, 0, 0)
        Else
            .Range("F" & rowIdx & ":H" & rowIdx).Font.Color = RGB(0, 128, 0)
        End If
        
        .Range("F" & rowIdx & ":H" & rowIdx).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range("F" & rowIdx & ":H" & rowIdx).Borders(xlEdgeBottom).Weight = xlThick
        
        .Range("A" & startRow & ":H" & startRow + 5).VerticalAlignment = xlCenter
    End With
End Sub

Private Sub WriteStatementFooter(ws As Worksheet, ByVal startRow As Long)
    WriteSharedFooter ws, startRow, False, "H"
End Sub

Private Sub SetupStatementPrint(ws As Worksheet, ByVal lastUsedRow As Long)
    On Error Resume Next
    
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$" & STMT_TABLE_HEADER
        .PrintArea = "$A$1:$H$" & lastUsedRow
        .CenterFooter = "Page &P of &N"
        .RightFooter = "Generated " & Format(Date, "YYYY/MM/DD")
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.75)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .Draft = False
        .BlackAndWhite = False
    End With
    
    PrepareStatementCenterHeaderForPrint ws
    
    On Error GoTo 0
End Sub

Private Sub PrepareStatementCenterHeaderForPrint(ws As Worksheet)
    Dim ps As PageSetup
    Dim s As String
    
    On Error GoTo SafeExit
    
    Set ps = ws.PageSetup
    ps.Draft = False
    
    s = ps.CenterHeader
    If InStr(1, s, "&G", vbTextCompare) > 0 Then
        ps.CenterHeader = Replace(s, "&G", "")
        ps.CenterHeader = s
    End If
    
SafeExit:
End Sub

Public Sub ExportStatementToPDF()
    Dim ws As Worksheet, savePath As String, fileName As String, custName As String
    
    If Not SheetExists(SHEET_STATEMENT) Then
        ShowError "Statement sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_STATEMENT)
    
    custName = CleanFileName(SafeString(ws.Range("B8").Value))
    If custName = "" Then custName = "Statement"
    fileName = custName & " - Statement " & Format(Date, "yyyy-mm") & ".pdf"
    
    savePath = GetSavePath(fileName, "PDF Files (*.pdf), *.pdf", "Save Statement as PDF")
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    SafeStatusBar "Exporting Statement PDF..."
    
    PrepareStatementCenterHeaderForPrint ws
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=savePath, Quality:=xlQualityStandard
    
    SafeStatusBar False
    ShowInfo "Statement PDF exported to:" & vbNewLine & vbNewLine & savePath, "Export Complete"
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error exporting Statement PDF: " & Err.Description
End Sub

Public Sub ExportStatementToExcel()
    ShowWarning "Export Statement to Excel is not yet implemented.", "Coming Soon"
End Sub

Public Sub PrintStatement()
    If Not SheetExists(SHEET_STATEMENT) Then
        ShowError "Statement sheet not found."
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    PrepareStatementCenterHeaderForPrint ThisWorkbook.Sheets(SHEET_STATEMENT)
    ThisWorkbook.Sheets(SHEET_STATEMENT).PrintOut Copies:=1, Preview:=True
    Exit Sub

ErrorHandler:
    ShowError "Error printing statement: " & Err.Description
End Sub

Public Function GetCustomerList() As Collection
    Dim col As New Collection
    Dim ws As Worksheet, lastRow As Long, i As Long, custName As String
    
    Set GetCustomerList = col
    
    If Not SheetExists(SHEET_CUSTOMERS) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(SHEET_CUSTOMERS)
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    On Error Resume Next
    For i = 2 To lastRow
        custName = Trim(CStr(ws.Cells(i, "A").Value))
        If custName <> "" Then
            col.Add custName, custName
        End If
    Next i
    On Error GoTo 0
    
    Set GetCustomerList = col
End Function

' =====================================================
' ===== FORM LAUNCHERS (explicit New / Unload lifecycle) =====
' =====================================================
' Using `New` (not the default instance) means each call gets a clean form,
' and `Unload` releases it cleanly. This avoids the VBA default-instance
' anti-pattern where stale form state can leak between launches.

Public Sub ShowStatementForm()
    Dim f As frmStatement
    Set f = New frmStatement
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub

Public Sub ShowMarkAsPaidForm()
    Dim f As frmMarkPaid
    Set f = New frmMarkPaid
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub

