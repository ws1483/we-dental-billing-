Attribute VB_Name = "modOpenQuotes"
Option Explicit

' =====================================================
' ===== OPEN QUOTES PANEL - RENDERS ON MENU SHEET =====
' Now with pagination (v1.5.0).
'
' Pagination model:
'   - Current page stored in MENU_OPEN_QUOTES_STATE_CELL (cell K28, hidden by font color)
'   - Page 1 = rows 1..MAX, Page 2 = rows (MAX+1)..(2*MAX), etc.
'   - Page resets to 1 whenever a filter changes (handled by the
'     Clear*Filter subs and by the Sheet10 double-click handlers).
'   - "Prev"/"Next" double-click cells are wired in Sheet10 code-behind
'     and call OpenQuotesNextPage / OpenQuotesPrevPage below.
' =====================================================

Public Sub RefreshOpenQuotesList()
    Dim wsMenu As Worksheet
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim trackingPath As String
    Dim lastTrackRow As Long, r As Long, outRow As Long
    Dim qNum As String, qDate As Variant, cust As String, patient As String
    Dim appType As String, status As String, totalIncl As Double
    Dim filePath As String
    Dim ageDays As Long
    Dim qualifyingCount As Long, pageCount As Long
    Dim totalSum As Double, pageSum As Double
    Dim includeRow As Boolean
    
    Dim hasDateFilter As Boolean, dFrom As Date, dTo As Date
    Dim hasCustFilter As Boolean, custFilter As String
    
    Dim currentPage As Long, totalPages As Long
    Dim pageStartIndex As Long, pageEndIndex As Long
    
    If Not SheetExists(SHEET_MENU) Then Exit Sub
    Set wsMenu = ThisWorkbook.Sheets(SHEET_MENU)
    
    GetOpenQuotesDateFilter hasDateFilter, dFrom, dTo
    GetOpenQuotesCustomerFilter hasCustFilter, custFilter
    
    currentPage = GetOpenQuotesCurrentPage()
    
    On Error Resume Next
    Application.CutCopyMode = False
    If ActiveSheet.Name = wsMenu.Name Then wsMenu.Range("A1").Select
    On Error GoTo 0
    
    ClearOpenQuotesPanel wsMenu
    UpdateOpenQuotesHeaderRange wsMenu, hasDateFilter, dFrom, dTo, hasCustFilter, custFilter
    
    trackingPath = GetQuoteTrackingPath()
    If trackingPath = "" Then
        WriteListEmptyMessage wsMenu, "Quote tracking file not found.", MENU_OPEN_QUOTES_FIRST_DATA
        RenderOpenQuotesPagerLabel wsMenu, 1, 1
        Exit Sub
    End If
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then
        WriteListEmptyMessage wsMenu, "Cannot open quote tracking file.", MENU_OPEN_QUOTES_FIRST_DATA
        RenderOpenQuotesPagerLabel wsMenu, 1, 1
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_QUOTES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        WriteListEmptyMessage wsMenu, "No quote tracking data found.", MENU_OPEN_QUOTES_FIRST_DATA
        RenderOpenQuotesPagerLabel wsMenu, 1, 1
        Exit Sub
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_QUOTES)
    lastTrackRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row
    
    ' ===== PASS 1: count qualifying rows (and sum totals for the grand total) =====
    qualifyingCount = 0
    totalSum = 0
    
    For r = 2 To lastTrackRow
        qNum = Trim(CStr(wsTrack.Cells(r, COL_Q_NUMBER).Value))
        status = Trim(CStr(wsTrack.Cells(r, COL_Q_STATUS).Value))
        
        includeRow = True
        
        If qNum = "" Then
            includeRow = False
        ElseIf LCase$(status) = LCase$(STATUS_QUOTE_CONVERTED) Then
            includeRow = False
        ElseIf LCase$(status) = LCase$(STATUS_QUOTE_CANCELLED) Then
            includeRow = False
        End If
        
        If includeRow Then
            qDate = wsTrack.Cells(r, COL_Q_DATE).Value
            cust = Trim(CStr(wsTrack.Cells(r, COL_Q_CUSTOMER).Value))
            
            If hasDateFilter Then
                If Not IsDate(qDate) Then
                    includeRow = False
                Else
                    If CDate(qDate) < dFrom Or CDate(qDate) > dTo Then includeRow = False
                End If
            End If
            
            If includeRow And hasCustFilter Then
                If StrComp(cust, custFilter, vbTextCompare) <> 0 Then includeRow = False
            End If
        End If
        
        If includeRow Then
            qualifyingCount = qualifyingCount + 1
            totalSum = totalSum + r2(wsTrack.Cells(r, COL_Q_TOTAL).Value)
        End If
    Next r
    
    ' ===== Compute pagination =====
    totalPages = (qualifyingCount + MENU_OPEN_QUOTES_MAX_ROWS - 1) \ MENU_OPEN_QUOTES_MAX_ROWS
    If totalPages < 1 Then totalPages = 1
    
    ' Clamp current page in case the underlying data shrank
    If currentPage > totalPages Then
        currentPage = totalPages
        SetOpenQuotesCurrentPage currentPage
    End If
    If currentPage < 1 Then
        currentPage = 1
        SetOpenQuotesCurrentPage currentPage
    End If
    
    pageStartIndex = (currentPage - 1) * MENU_OPEN_QUOTES_MAX_ROWS + 1   ' 1-based
    pageEndIndex = pageStartIndex + MENU_OPEN_QUOTES_MAX_ROWS - 1
    
    ' ===== PASS 2: render rows that fall in the current page window =====
    outRow = MENU_OPEN_QUOTES_FIRST_DATA
    pageCount = 0
    pageSum = 0
    
    Dim qualifyingIdx As Long
    qualifyingIdx = 0
    
    For r = 2 To lastTrackRow
        qNum = Trim(CStr(wsTrack.Cells(r, COL_Q_NUMBER).Value))
        status = Trim(CStr(wsTrack.Cells(r, COL_Q_STATUS).Value))
        
        includeRow = True
        
        If qNum = "" Then
            includeRow = False
        ElseIf LCase$(status) = LCase$(STATUS_QUOTE_CONVERTED) Then
            includeRow = False
        ElseIf LCase$(status) = LCase$(STATUS_QUOTE_CANCELLED) Then
            includeRow = False
        End If
        
        If includeRow Then
            qDate = wsTrack.Cells(r, COL_Q_DATE).Value
            cust = Trim(CStr(wsTrack.Cells(r, COL_Q_CUSTOMER).Value))
            
            If hasDateFilter Then
                If Not IsDate(qDate) Then
                    includeRow = False
                Else
                    If CDate(qDate) < dFrom Or CDate(qDate) > dTo Then includeRow = False
                End If
            End If
            
            If includeRow And hasCustFilter Then
                If StrComp(cust, custFilter, vbTextCompare) <> 0 Then includeRow = False
            End If
        End If
        
        If includeRow Then
            qualifyingIdx = qualifyingIdx + 1
            
            ' Only render rows in [pageStartIndex, pageEndIndex]
            If qualifyingIdx >= pageStartIndex And qualifyingIdx <= pageEndIndex Then
                patient = Trim(CStr(wsTrack.Cells(r, COL_Q_PATIENT).Value))
                appType = Trim(CStr(wsTrack.Cells(r, COL_Q_APPLIANCE).Value))
                totalIncl = r2(wsTrack.Cells(r, COL_Q_TOTAL).Value)
                filePath = Trim(CStr(wsTrack.Cells(r, COL_Q_FILEPATH).Value))
                
                ageDays = 0
                If IsDate(qDate) Then
                    ageDays = DateDiff("d", CDate(qDate), Date)
                    If ageDays < 0 Then ageDays = 0
                End If
                
                WriteOpenQuoteRow wsMenu, outRow, qDate, qNum, cust, patient, appType, _
                                  totalIncl, ageDays, filePath
                
                pageSum = pageSum + totalIncl
                pageCount = pageCount + 1
                outRow = outRow + 1
            ElseIf qualifyingIdx > pageEndIndex Then
                Exit For   ' nothing left to render on this page
            End If
        End If
    Next r
    
    CloseTrackingWorkbook wbTrack, False
    
    ' ===== Render summary / empty state =====
    If qualifyingCount = 0 Then
        WriteListEmptyMessage wsMenu, _
            BuildListEmptyMessage(hasDateFilter, dFrom, dTo, hasCustFilter, custFilter, True), _
            MENU_OPEN_QUOTES_FIRST_DATA
    Else
        WriteOpenQuotesTotal wsMenu, MENU_OPEN_QUOTES_TOTAL_ROW, qualifyingCount, totalSum, _
                             pageCount, pageSum, currentPage, totalPages
        WriteListHintLine wsMenu, MENU_OPEN_QUOTES_TIP_ROW, _
            "Tip: Double-click a blue Quote # to open; double-click PREV / NEXT cells to page"
    End If
    
    RenderOpenQuotesPagerLabel wsMenu, currentPage, totalPages
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    WriteListEmptyMessage wsMenu, "Error: " & Err.Description, MENU_OPEN_QUOTES_FIRST_DATA
    RenderOpenQuotesPagerLabel wsMenu, 1, 1
End Sub

' =====================================================
' ===== PAGINATION PUBLIC API (called from Sheet10) =====
' =====================================================

Public Sub OpenQuotesNextPage()
    Dim cur As Long
    cur = GetOpenQuotesCurrentPage()
    SetOpenQuotesCurrentPage cur + 1
    RefreshOpenQuotesList    ' will clamp if past end
End Sub

Public Sub OpenQuotesPrevPage()
    Dim cur As Long
    cur = GetOpenQuotesCurrentPage()
    cur = cur - 1
    If cur < 1 Then cur = 1
    SetOpenQuotesCurrentPage cur
    RefreshOpenQuotesList
End Sub

Public Sub OpenQuotesResetToFirstPage()
    SetOpenQuotesCurrentPage 1
End Sub

' =====================================================
' ===== PAGINATION STATE (stored in hidden cell) =====
' =====================================================

Private Function GetOpenQuotesCurrentPage() As Long
    Dim ws As Worksheet, v As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then
        GetOpenQuotesCurrentPage = 1
        Exit Function
    End If
    v = ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Value
    If IsNumeric(v) Then
        If CLng(v) >= 1 Then
            GetOpenQuotesCurrentPage = CLng(v)
            Exit Function
        End If
    End If
    GetOpenQuotesCurrentPage = 1
    On Error GoTo 0
End Function

Private Sub SetOpenQuotesCurrentPage(ByVal pageNum As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then Exit Sub
    If pageNum < 1 Then pageNum = 1
    ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Value = pageNum
    ' Keep the state cell hidden by matching font to background
    ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Font.Color = ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Interior.Color
    On Error GoTo 0
End Sub

Private Sub RenderOpenQuotesPagerLabel(ws As Worksheet, _
                                       ByVal currentPage As Long, _
                                       ByVal totalPages As Long)
    On Error Resume Next
    
    If totalPages < 1 Then totalPages = 1
    If currentPage < 1 Then currentPage = 1
    If currentPage > totalPages Then currentPage = totalPages
    
    ' --- PREV cell ---
    With ws.Range(MENU_OPEN_QUOTES_PREV_CELL)
        .Value = ChrW(9664) & " Prev"   ' ? Prev
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
    
    ' --- PAGE label ---
    With ws.Range(MENU_OPEN_QUOTES_PAGE_CELL)
        .Value = "Page " & currentPage & " of " & totalPages
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.bold = True
        .Font.Color = RGB(46, 135, 176)
        .Interior.Pattern = xlNone
    End With
    
    ' --- NEXT cell ---
    With ws.Range(MENU_OPEN_QUOTES_NEXT_CELL)
        .Value = "Next " & ChrW(9654)   ' Next ?
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

Public Sub ClearOpenQuotesCustomerFilter()
    On Error Resume Next
    If Not SheetExists(SHEET_MENU) Then Exit Sub
    ThisWorkbook.Sheets(SHEET_MENU).Range(MENU_OPEN_QUOTES_CUST_CELL).ClearContents
    OpenQuotesResetToFirstPage   ' filter changed -> back to page 1
    RefreshOpenQuotesList
    On Error GoTo 0
End Sub
' =====================================================
' ===== OPEN QUOTES DATE FILTER (was missing) =====
' =====================================================
' Reads optional From/To date cells (D27 / F27) on the Menu sheet.
' Open Quotes panel doesn't currently expose these cells in the
' SetupMenuPanels layout, so by default this returns hasFilter = False.
' If you ever add date cells to row 27, this function will start
' filtering automatically.

Private Sub GetOpenQuotesDateFilter(ByRef hasFilter As Boolean, _
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
    
    vFrom = ws.Range("D27").Value
    vTo = ws.Range("F27").Value
    
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
Private Sub GetOpenQuotesCustomerFilter(ByRef hasFilter As Boolean, ByRef custName As String)
    Dim ws As Worksheet
    hasFilter = False
    custName = ""
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    If ws Is Nothing Then Exit Sub
    custName = Trim(CStr(ws.Range(MENU_OPEN_QUOTES_CUST_CELL).Value))
    If custName <> "" Then hasFilter = True
    On Error GoTo 0
End Sub

Private Sub UpdateOpenQuotesHeaderRange(ws As Worksheet, _
                                        ByVal hasDateFilter As Boolean, _
                                        ByVal dFrom As Date, _
                                        ByVal dTo As Date, _
                                        ByVal hasCustFilter As Boolean, _
                                        ByVal custFilter As String)
    Dim hrLabelRow As Long, txt As String
    hrLabelRow = MENU_OPEN_QUOTES_HEADER_ROW - 1   ' title bar row
    
    On Error Resume Next
    
    txt = "Open Quotes"
    If hasCustFilter Then txt = txt & "  -  " & custFilter
    If hasDateFilter Then txt = txt & "  -  " & Format(dFrom, "YYYY/MM/DD") & " to " & Format(dTo, "YYYY/MM/DD")
    If Not hasDateFilter And Not hasCustFilter Then txt = txt & " (not yet converted to invoice)"
    
    ws.Range("B" & hrLabelRow & ":E" & hrLabelRow).UnMerge
    ws.Range("B" & hrLabelRow & ":E" & hrLabelRow).ClearContents
    
    With ws.Range("B" & hrLabelRow & ":E" & hrLabelRow)
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

' Clears only the DATA area: from first data row down through summary + tip lines.
' Uses derived constants so we can never accidentally overrun into the Unpaid panel.
' Does NOT clear pager cells (those are written by RenderOpenQuotesPagerLabel).
Private Sub ClearOpenQuotesPanel(ws As Worksheet)
    Dim r1 As Long, r2 As Long
    
    r1 = MENU_OPEN_QUOTES_FIRST_DATA
    r2 = MENU_OPEN_QUOTES_TIP_ROW          ' includes totals row + tip row
    
    ' Safety net: never bleed into the Unpaid Invoices panel above row 49.
    If r2 >= MENU_UNPAID_INV_TITLE_ROW Then r2 = MENU_UNPAID_INV_TITLE_ROW - 1
    
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

' --- Shared helpers (also used by modUnpaidInvoices) ---

Public Sub WriteListEmptyMessage(ws As Worksheet, ByVal message As String, ByVal r As Long)
    On Error Resume Next
    ws.Range("B" & r & ":H" & r).UnMerge
    With ws.Range("B" & r & ":H" & r)
        .Merge
        .Value = message
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
        .HorizontalAlignment = xlCenter
    End With
    On Error GoTo 0
End Sub

Public Function BuildListEmptyMessage(ByVal hasDateFilter As Boolean, _
                                       ByVal dFrom As Date, _
                                       ByVal dTo As Date, _
                                       ByVal hasCustFilter As Boolean, _
                                       ByVal custFilter As String, _
                                       ByVal isQuotes As Boolean) As String
    Dim s As String
    If isQuotes Then
        s = "No open quotes"
    Else
        s = "No unpaid invoices"
    End If
    If hasCustFilter Then s = s & " for " & custFilter
    If hasDateFilter Then s = s & " between " & Format(dFrom, "YYYY/MM/DD") & " and " & Format(dTo, "YYYY/MM/DD")
    If Not hasDateFilter And Not hasCustFilter Then
        If isQuotes Then
            s = s & " - everything is converted or cancelled."
        Else
            s = s & " - all invoices are paid."
        End If
    Else
        s = s & "."
    End If
    BuildListEmptyMessage = s
End Function

Public Sub WriteListHintLine(ws As Worksheet, ByVal outRow As Long, ByVal text As String)
    On Error Resume Next
    ws.Range("B" & outRow & ":H" & outRow).UnMerge
    With ws.Range("B" & outRow & ":H" & outRow)
        .Merge
        .Value = text
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)
        .HorizontalAlignment = xlLeft
    End With
    On Error GoTo 0
End Sub

' --- Row writers ---

Private Sub WriteOpenQuoteRow(ws As Worksheet, ByVal outRow As Long, _
                              ByVal qDate As Variant, ByVal qNum As String, _
                              ByVal cust As String, ByVal patient As String, _
                              ByVal appType As String, ByVal totalIncl As Double, _
                              ByVal ageDays As Long, ByVal filePath As String)
    With ws
        If IsDate(qDate) Then
            .Cells(outRow, "B").Value = CDate(qDate)
            .Cells(outRow, "B").NumberFormat = "YYYY/MM/DD"
        Else
            .Cells(outRow, "B").Value = SafeString(qDate)
        End If
        
        .Cells(outRow, "C").Value = qNum
        .Cells(outRow, "C").Font.Color = RGB(46, 135, 176)
        .Cells(outRow, "C").Font.bold = True
        .Cells(outRow, "C").Font.Underline = xlUnderlineStyleSingle
        
        .Cells(outRow, "D").Value = cust
        .Cells(outRow, "E").Value = patient
        .Cells(outRow, "F").Value = appType
        
        .Cells(outRow, "G").Value = Round(totalIncl, 2)
        .Cells(outRow, "G").NumberFormat = """R""#,##0.00"
        .Cells(outRow, "G").HorizontalAlignment = xlRight
        
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

' Now takes (qualifyingCount, totalSum) AND (pageCount, pageSum, page, totalPages)
' so the totals row can show both "Total across all matching" AND "On this page".
Private Sub WriteOpenQuotesTotal(ws As Worksheet, ByVal outRow As Long, _
                                 ByVal qualifyingCount As Long, ByVal totalSum As Double, _
                                 ByVal pageCount As Long, ByVal pageSum As Double, _
                                 ByVal currentPage As Long, ByVal totalPages As Long)
    Dim lineText As String
    
    With ws
        .Range("B" & outRow & ":H" & outRow).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("B" & outRow & ":H" & outRow).Borders(xlEdgeTop).Weight = xlMedium
        .Range("B" & outRow & ":H" & outRow).Borders(xlEdgeTop).Color = RGB(46, 135, 176)
        
        If totalPages > 1 Then
            lineText = pageCount & " on page  |  " & qualifyingCount & " total open quote(s)"
        Else
            lineText = qualifyingCount & " open quote(s)"
        End If
        
        .Cells(outRow, "B").Value = lineText
        .Cells(outRow, "B").Font.bold = True
        .Cells(outRow, "B").Font.Color = RGB(46, 135, 176)
        
        .Cells(outRow, "F").Value = "Total:"
        .Cells(outRow, "F").Font.bold = True
        .Cells(outRow, "F").HorizontalAlignment = xlRight
        
        .Cells(outRow, "G").Value = Round(totalSum, 2)
        .Cells(outRow, "G").NumberFormat = """R""#,##0.00"
        .Cells(outRow, "G").Font.bold = True
        .Cells(outRow, "G").HorizontalAlignment = xlRight
        
        .Range("B" & outRow & ":H" & outRow).VerticalAlignment = xlCenter
    End With
End Sub

