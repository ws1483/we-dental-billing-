Attribute VB_Name = "modUnpaidInvoices"
Option Explicit

' =====================================================
' ===== UNPAID INVOICES PANEL (Open + Partial) =====
' Renders below the Open Quotes panel on the Menu sheet.
' Now with pagination (v1.5.0).
'
' Pagination model: same as Open Quotes (state cell, paging API,
' clamp when underlying data shrinks).
' =====================================================

Public Sub RefreshUnpaidInvoicesList()
    Dim wsMenu As Worksheet
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim trackingPath As String
    Dim lastTrackRow As Long, r As Long, outRow As Long
    Dim invNum As String, iDate As Variant, cust As String, patient As String
    Dim appType As String, status As String, totalIncl As Double
    Dim paidAmt As Double, creditAmt As Double, balance As Double
    Dim filePath As String
    Dim ageDays As Long
    Dim qualifyingCount As Long, pageCount As Long
    Dim totalSum As Double, pageSum As Double
    Dim includeRow As Boolean
    
    Dim hasDateFilter As Boolean, dFrom As Date, dTo As Date
    Dim hasCustFilter As Boolean, custFilter As String
    
    Dim currentPage As Long, totalPages As Long
    Dim pageStartIndex As Long, pageEndIndex As Long
    Dim qualifyingIdx As Long
    
    If Not SheetExists(SHEET_MENU) Then Exit Sub
    Set wsMenu = ThisWorkbook.Sheets(SHEET_MENU)
    
    GetUnpaidInvoicesDateFilter hasDateFilter, dFrom, dTo
    GetUnpaidInvoicesCustomerFilter hasCustFilter, custFilter
    
    currentPage = GetUnpaidInvoicesCurrentPage()
    
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
    
    ClearUnpaidInvoicesPanel wsMenu
    UpdateUnpaidInvoicesHeaderRange wsMenu, hasDateFilter, dFrom, dTo, hasCustFilter, custFilter
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then
        WriteListEmptyMessage wsMenu, "Invoice tracking file not found.", MENU_UNPAID_INV_FIRST_DATA
        RenderUnpaidInvoicesPagerLabel wsMenu, 1, 1
        Exit Sub
    End If
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then
        WriteListEmptyMessage wsMenu, "Cannot open invoice tracking file.", MENU_UNPAID_INV_FIRST_DATA
        RenderUnpaidInvoicesPagerLabel wsMenu, 1, 1
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        WriteListEmptyMessage wsMenu, "No invoice tracking data found.", MENU_UNPAID_INV_FIRST_DATA
        RenderUnpaidInvoicesPagerLabel wsMenu, 1, 1
        Exit Sub
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    lastTrackRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row
    
    ' ===== PASS 1: count qualifying rows and sum outstanding balance =====
    qualifyingCount = 0
    totalSum = 0
    
    For r = 2 To lastTrackRow
        invNum = Trim(CStr(wsTrack.Cells(r, COL_I_NUMBER).Value))
        status = Trim(CStr(wsTrack.Cells(r, COL_I_STATUS).Value))
        
        includeRow = True
        
        If invNum = "" Then
            includeRow = False
        ElseIf Not IsUnpaidStatus(status) Then
            includeRow = False
        End If
        
        If includeRow Then
            iDate = wsTrack.Cells(r, COL_I_DATE).Value
            cust = Trim(CStr(wsTrack.Cells(r, COL_I_CUSTOMER).Value))
            
            If hasDateFilter Then
                If Not IsDate(iDate) Then
                    includeRow = False
                Else
                    If CDate(iDate) < dFrom Or CDate(iDate) > dTo Then includeRow = False
                End If
            End If
            
            If includeRow And hasCustFilter Then
                If StrComp(cust, custFilter, vbTextCompare) <> 0 Then includeRow = False
            End If
        End If
        
        If includeRow Then
            totalIncl = r2(wsTrack.Cells(r, COL_I_TOTAL).Value)
            paidAmt = r2(wsTrack.Cells(r, COL_I_PAIDAMOUNT).Value)
            creditAmt = r2(wsTrack.Cells(r, COL_I_CREDITAMT).Value)
            
            balance = Round(totalIncl - paidAmt - creditAmt, 2)
            If balance < 0 Then balance = 0
            
            If balance <= 0 Then includeRow = False
        End If
        
        If includeRow Then
            qualifyingCount = qualifyingCount + 1
            totalSum = totalSum + balance
        End If
    Next r
    
    ' ===== Compute pagination =====
    totalPages = (qualifyingCount + MENU_UNPAID_INV_MAX_ROWS - 1) \ MENU_UNPAID_INV_MAX_ROWS
    If totalPages < 1 Then totalPages = 1
    
    If currentPage > totalPages Then
        currentPage = totalPages
        SetUnpaidInvoicesCurrentPage currentPage
    End If
    If currentPage < 1 Then
        currentPage = 1
        SetUnpaidInvoicesCurrentPage currentPage
    End If
    
    pageStartIndex = (currentPage - 1) * MENU_UNPAID_INV_MAX_ROWS + 1
    pageEndIndex = pageStartIndex + MENU_UNPAID_INV_MAX_ROWS - 1
    
    ' ===== PASS 2: render only rows on current page =====
    outRow = MENU_UNPAID_INV_FIRST_DATA
    pageCount = 0
    pageSum = 0
    qualifyingIdx = 0
    
    For r = 2 To lastTrackRow
        invNum = Trim(CStr(wsTrack.Cells(r, COL_I_NUMBER).Value))
        status = Trim(CStr(wsTrack.Cells(r, COL_I_STATUS).Value))
        
        includeRow = True
        
        If invNum = "" Then
            includeRow = False
        ElseIf Not IsUnpaidStatus(status) Then
            includeRow = False
        End If
        
        If includeRow Then
            iDate = wsTrack.Cells(r, COL_I_DATE).Value
            cust = Trim(CStr(wsTrack.Cells(r, COL_I_CUSTOMER).Value))
            
            If hasDateFilter Then
                If Not IsDate(iDate) Then
                    includeRow = False
                Else
                    If CDate(iDate) < dFrom Or CDate(iDate) > dTo Then includeRow = False
                End If
            End If
            
            If includeRow And hasCustFilter Then
                If StrComp(cust, custFilter, vbTextCompare) <> 0 Then includeRow = False
            End If
        End If
        
        If includeRow Then
            patient = Trim(CStr(wsTrack.Cells(r, COL_I_PATIENT).Value))
            appType = Trim(CStr(wsTrack.Cells(r, COL_I_APPLIANCE).Value))
            totalIncl = r2(wsTrack.Cells(r, COL_I_TOTAL).Value)
            paidAmt = r2(wsTrack.Cells(r, COL_I_PAIDAMOUNT).Value)
            creditAmt = r2(wsTrack.Cells(r, COL_I_CREDITAMT).Value)
            filePath = Trim(CStr(wsTrack.Cells(r, COL_I_FILEPATH).Value))
            
            balance = Round(totalIncl - paidAmt - creditAmt, 2)
            If balance < 0 Then balance = 0
            
            If balance <= 0 Then includeRow = False
        End If
        
        If includeRow Then
            qualifyingIdx = qualifyingIdx + 1
            
            If qualifyingIdx >= pageStartIndex And qualifyingIdx <= pageEndIndex Then
                ageDays = 0
                If IsDate(iDate) Then
                    ageDays = DateDiff("d", CDate(iDate), Date)
                    If ageDays < 0 Then ageDays = 0
                End If
                
                WriteUnpaidInvoiceRow wsMenu, outRow, iDate, invNum, cust, patient, appType, _
                                      balance, ageDays, filePath, status
                
                pageSum = pageSum + balance
                pageCount = pageCount + 1
                outRow = outRow + 1
            ElseIf qualifyingIdx > pageEndIndex Then
                Exit For
            End If
        End If
    Next r
    
    CloseTrackingWorkbook wbTrack, False
    
    If qualifyingCount = 0 Then
        WriteListEmptyMessage wsMenu, _
            BuildListEmptyMessage(hasDateFilter, dFrom, dTo, hasCustFilter, custFilter, False), _
            MENU_UNPAID_INV_FIRST_DATA
    Else
        Dim totalRow As Long, tipRow As Long
        totalRow = MENU_UNPAID_INV_FIRST_DATA + MENU_UNPAID_INV_MAX_ROWS
        tipRow = totalRow + 1
        
        WriteUnpaidInvoicesTotal wsMenu, totalRow, qualifyingCount, totalSum, _
                                 pageCount, pageSum, currentPage, totalPages
        WriteListHintLine wsMenu, tipRow, _
            "Tip: Double-click a blue Invoice # to open; double-click PREV / NEXT cells to page"
    End If
    
    RenderUnpaidInvoicesPagerLabel wsMenu, currentPage, totalPages
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    WriteListEmptyMessage wsMenu, "Error: " & Err.Description, MENU_UNPAID_INV_FIRST_DATA
    RenderUnpaidInvoicesPagerLabel wsMenu, 1, 1
End Sub

' =====================================================
' ===== PAGINATION PUBLIC API (called from Sheet10) =====
' =====================================================

Public Sub UnpaidInvoicesNextPage()
    Dim cur As Long
    cur = GetUnpaidInvoicesCurrentPage()
    SetUnpaidInvoicesCurrentPage cur + 1
    RefreshUnpaidInvoicesList
End Sub

Public Sub UnpaidInvoicesPrevPage()
    Dim cur As Long
    cur = GetUnpaidInvoicesCurrentPage()
    cur = cur - 1
    If cur < 1 Then cur = 1
    SetUnpaidInvoicesCurrentPage cur
    RefreshUnpaidInvoicesList
End Sub

Public Sub UnpaidInvoicesResetToFirstPage()
    SetUnpaidInvoicesCurrentPage 1
End Sub

' =====================================================
' ===== PAGINATION STATE (stored in hidden cell) =====
' =====================================================

Private Function GetUnpaidInvoicesCurrentPage() As Long
    Dim ws As Worksheet, v As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then
        GetUnpaidInvoicesCurrentPage = 1
        Exit Function
    End If
    v = ws.Range(MENU_UNPAID_INV_STATE_CELL).Value
    If IsNumeric(v) Then
        If CLng(v) >= 1 Then
            GetUnpaidInvoicesCurrentPage = CLng(v)
            Exit Function
        End If
    End If
    GetUnpaidInvoicesCurrentPage = 1
    On Error GoTo 0
End Function

Private Sub SetUnpaidInvoicesCurrentPage(ByVal pageNum As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then Exit Sub
    If pageNum < 1 Then pageNum = 1
    ws.Range(MENU_UNPAID_INV_STATE_CELL).Value = pageNum
    ws.Range(MENU_UNPAID_INV_STATE_CELL).Font.Color = _
        ws.Range(MENU_UNPAID_INV_STATE_CELL).Interior.Color
    On Error GoTo 0
End Sub

Private Sub RenderUnpaidInvoicesPagerLabel(ws As Worksheet, _
                                           ByVal currentPage As Long, _
                                           ByVal totalPages As Long)
    On Error Resume Next
    
    If totalPages < 1 Then totalPages = 1
    If currentPage < 1 Then currentPage = 1
    If currentPage > totalPages Then currentPage = totalPages
    
    With ws.Range(MENU_UNPAID_INV_PREV_CELL)
        .Value = ChrW(9664) & " Prev"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.bold = True
        If currentPage <= 1 Then
            .Font.Color = RGB(180, 180, 180)
            .Interior.Color = RGB(245, 245, 245)
        Else
            .Font.Color = RGB(46, 135, 176)
            .Interior.Color = RGB(225, 240, 250)
        End If
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 200, 220)
    End With
    
    With ws.Range(MENU_UNPAID_INV_PAGE_CELL)
        .Value = "Page " & currentPage & " of " & totalPages
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.bold = True
        .Font.Color = RGB(255, 255, 255)
        ' The header row 51 has a blue background (set by SetupUnpaidInvoicesPanel)
        ' so white text reads well there.
    End With
    
    With ws.Range(MENU_UNPAID_INV_NEXT_CELL)
        .Value = "Next " & ChrW(9654)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.bold = True
        If currentPage >= totalPages Then
            .Font.Color = RGB(180, 180, 180)
            .Interior.Color = RGB(245, 245, 245)
        Else
            .Font.Color = RGB(46, 135, 176)
            .Interior.Color = RGB(225, 240, 250)
        End If
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 200, 220)
    End With
    
    On Error GoTo 0
End Sub

' =====================================================
' ===== FILTER MGMT =====
' =====================================================

Public Sub ClearUnpaidInvoicesCustomerFilter()
    On Error Resume Next
    If Not SheetExists(SHEET_MENU) Then Exit Sub
    ThisWorkbook.Sheets(SHEET_MENU).Range(MENU_UNPAID_INV_CUST_CELL).ClearContents
    UnpaidInvoicesResetToFirstPage
    RefreshUnpaidInvoicesList
    On Error GoTo 0
End Sub

' =====================================================
' ===== HELPERS =====
' =====================================================

Private Function IsUnpaidStatus(ByVal status As String) As Boolean
    Dim s As String
    s = LCase$(Trim(status))
    
    Select Case s
        Case ""
            IsUnpaidStatus = False
        Case "paid", "cancelled", "canceled", "void", "voided", "credit", "credited", "refunded"
            IsUnpaidStatus = False
        Case Else
            IsUnpaidStatus = True
    End Select
End Function
' =====================================================
' ===== UNPAID INVOICES DATE FILTER (was missing) =====
' =====================================================
' Reads the From/To date cells on the Menu sheet.
' These cells already exist (C49 and F49 from SetupUnpaidInvoicesPanel),
' so this filter is fully wired and will start working immediately.

Private Sub GetUnpaidInvoicesDateFilter(ByRef hasFilter As Boolean, _
                                         ByRef dFrom As Date, _
                                         ByRef dTo As Date)
    Dim ws As Worksheet
    Dim vFrom As Variant, vTo As Variant
    Dim tmp As Date
    
    hasFilter = False
    dFrom = 0
    dTo = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then Exit Sub
    
    vFrom = ws.Range("C49").Value
    vTo = ws.Range("F49").Value
    
    If IsDate(vFrom) Then dFrom = CDate(vFrom)
    If IsDate(vTo) Then dTo = CDate(vTo)
    
    If dFrom > 0 And dTo > 0 Then
        hasFilter = True
        If dFrom > dTo Then
            tmp = dFrom
            dFrom = dTo
            dTo = tmp
        End If
    ElseIf dFrom > 0 And dTo = 0 Then
        hasFilter = True
        dTo = DateSerial(9999, 12, 31)
    ElseIf dFrom = 0 And dTo > 0 Then
        hasFilter = True
        dFrom = DateSerial(1900, 1, 1)
    End If
    
    On Error GoTo 0
End Sub
Private Sub GetUnpaidInvoicesCustomerFilter(ByRef hasFilter As Boolean, ByRef custName As String)
    Dim ws As Worksheet
    hasFilter = False
    custName = ""
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then Exit Sub
    custName = Trim(CStr(ws.Range(MENU_UNPAID_INV_CUST_CELL).Value))
    If custName <> "" Then hasFilter = True
    On Error GoTo 0
End Sub

Private Sub UpdateUnpaidInvoicesHeaderRange(ws As Worksheet, _
                                            ByVal hasDateFilter As Boolean, _
                                            ByVal dFrom As Date, _
                                            ByVal dTo As Date, _
                                            ByVal hasCustFilter As Boolean, _
                                            ByVal custFilter As String)
    Dim titleRow As Long, txt As String
    titleRow = MENU_UNPAID_INV_TITLE_ROW
    
    On Error Resume Next
    
    txt = "Unpaid Invoices"
    If hasCustFilter Then txt = txt & "  -  " & custFilter
    If hasDateFilter Then txt = txt & "  -  " & Format(dFrom, "YYYY/MM/DD") & " to " & Format(dTo, "YYYY/MM/DD")
    If Not hasDateFilter And Not hasCustFilter Then txt = txt & " (Open + Partial)"
    
    ws.Range("B" & titleRow & ":E" & titleRow).UnMerge
    ws.Range("B" & titleRow & ":E" & titleRow).ClearContents
    
    With ws.Range("B" & titleRow & ":E" & titleRow)
        .Merge
        .Value = txt
        .Font.bold = True
        .Font.Size = 12
        .Font.Color = RGB(46, 135, 176)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo 0
End Sub

Private Sub ClearUnpaidInvoicesPanel(ws As Worksheet)
    Dim r1 As Long, r2 As Long
    
    ' Only clear the DATA area: from first data row down through the
    ' summary + tip lines. DO NOT touch rows 49-52 (title, filter, customer,
    ' header) - those are part of the static panel layout.
    r1 = MENU_UNPAID_INV_FIRST_DATA
    r2 = r1 + MENU_UNPAID_INV_MAX_ROWS + 2   ' +2 = summary line + tip line
    
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range("B" & r1 & ":I" & r2)
    rng.UnMerge
    rng.ClearComments
    rng.ClearContents
    rng.Interior.Pattern = xlNone
    rng.Borders.LineStyle = xlNone
    rng.Font.Color = vbBlack
    rng.Font.bold = False
    rng.Font.Italic = False
    rng.Font.Underline = xlUnderlineStyleNone
    On Error GoTo 0
End Sub

Private Sub WriteUnpaidInvoiceRow(ws As Worksheet, ByVal outRow As Long, _
                                  ByVal iDate As Variant, ByVal invNum As String, _
                                  ByVal cust As String, ByVal patient As String, _
                                  ByVal appType As String, ByVal balance As Double, _
                                  ByVal ageDays As Long, ByVal filePath As String, _
                                  ByVal status As String)
    With ws
        If IsDate(iDate) Then
            .Cells(outRow, "B").Value = CDate(iDate)
            .Cells(outRow, "B").NumberFormat = "YYYY/MM/DD"
        Else
            .Cells(outRow, "B").Value = SafeString(iDate)
        End If
        
        .Cells(outRow, "C").Value = invNum
        .Cells(outRow, "C").Font.Color = RGB(46, 135, 176)
        .Cells(outRow, "C").Font.bold = True
        .Cells(outRow, "C").Font.Underline = xlUnderlineStyleSingle
        
        .Cells(outRow, "D").Value = cust
        .Cells(outRow, "E").Value = patient
        .Cells(outRow, "F").Value = appType
        
        .Cells(outRow, "G").Value = Round(balance, 2)
        .Cells(outRow, "G").NumberFormat = """R""#,##0.00"
        .Cells(outRow, "G").HorizontalAlignment = xlRight
        
        If LCase$(status) = LCase$(STATUS_INVOICE_PARTIAL) Then
            .Cells(outRow, "G").Font.Color = RGB(204, 102, 0)
            .Cells(outRow, "G").Font.bold = True
        End If
        
        .Cells(outRow, "H").Value = ageDays & "d"
        .Cells(outRow, "H").HorizontalAlignment = xlCenter
        
        .Cells(outRow, "I").Value = filePath
        
        Select Case ageDays
            Case Is <= 7
                .Cells(outRow, "H").Font.Color = RGB(0, 128, 0)
            Case 8 To 30
                .Cells(outRow, "H").Font.Color = RGB(204, 153, 0)
            Case 31 To 60
                .Cells(outRow, "H").Font.Color = RGB(230, 100, 0)
            Case Else
                .Cells(outRow, "H").Font.Color = RGB(192, 0, 0)
                .Cells(outRow, "H").Font.bold = True
        End Select
        
        If (outRow Mod 2) = 0 Then
            .Range("B" & outRow & ":H" & outRow).Interior.Color = RGB(245, 248, 252)
        End If
        
        .Range("B" & outRow & ":H" & outRow).VerticalAlignment = xlCenter
        
        If filePath <> "" Then
            On Error Resume Next
            Dim cellC As Range
            Set cellC = .Cells(outRow, "C")
            If Not cellC Is Nothing Then
                If Not cellC.MergeCells Then
                    If Not cellC.Comment Is Nothing Then cellC.Comment.Delete
                    cellC.AddComment "Double-click to open: " & filePath
                    If Not cellC.Comment Is Nothing Then
                        cellC.Comment.Shape.TextFrame.AutoSize = True
                    End If
                End If
            End If
            On Error GoTo 0
        End If
    End With
End Sub

Private Sub WriteUnpaidInvoicesTotal(ws As Worksheet, ByVal outRow As Long, _
                                     ByVal qualifyingCount As Long, ByVal totalSum As Double, _
                                     ByVal pageCount As Long, ByVal pageSum As Double, _
                                     ByVal currentPage As Long, ByVal totalPages As Long)
    Dim lineText As String
    
    With ws
        .Range("B" & outRow & ":H" & outRow).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("B" & outRow & ":H" & outRow).Borders(xlEdgeTop).Weight = xlMedium
        .Range("B" & outRow & ":H" & outRow).Borders(xlEdgeTop).Color = RGB(46, 135, 176)
        
        If totalPages > 1 Then
            lineText = pageCount & " on page  |  " & qualifyingCount & " total unpaid invoice(s)"
        Else
            lineText = qualifyingCount & " unpaid invoice(s)"
        End If
        
        .Cells(outRow, "B").Value = lineText
        .Cells(outRow, "B").Font.bold = True
        .Cells(outRow, "B").Font.Color = RGB(46, 135, 176)
        
        .Cells(outRow, "F").Value = "Outstanding:"
        .Cells(outRow, "F").Font.bold = True
        .Cells(outRow, "F").HorizontalAlignment = xlRight
        
        .Cells(outRow, "G").Value = Round(totalSum, 2)
        .Cells(outRow, "G").NumberFormat = """R""#,##0.00"
        .Cells(outRow, "G").Font.bold = True
        .Cells(outRow, "G").Font.Color = RGB(192, 0, 0)
        .Cells(outRow, "G").HorizontalAlignment = xlRight
        
        .Range("B" & outRow & ":H" & outRow).VerticalAlignment = xlCenter
    End With
End Sub

