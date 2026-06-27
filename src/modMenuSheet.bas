Attribute VB_Name = "modMenuSheet"
Option Explicit

' =====================================================
' ===== ONE-TIME SETUP FOR MENU SHEET PANELS =====
' Run SetupMenuPanels ONCE (Alt+F8) to:
'   - add Customer dropdown row to Open Quotes panel
'   - draw the Unpaid Invoices panel below it
'   - set up Data Validation lists from Customers sheet
'   - initialise pagination state cells and prep their tooltips
' Safe to re-run — it overwrites cells in-place.
' =====================================================

Public Sub SetupMenuPanels()
    Dim ws As Worksheet
    
    If Not SheetExists(SHEET_MENU) Then
        MsgBox "Menu sheet '" & SHEET_MENU & "' not found.", vbCritical
        Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets(SHEET_MENU)
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    SetupOpenQuotesCustomerRow ws
    SetupUnpaidInvoicesPanel ws
    SetupCustomerValidation ws
    SetupPaginationCells ws        ' NEW
    
    Application.ScreenUpdating = True
    MsgBox "Menu panels set up successfully." & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Right-click the 'Clear' button next to the Open Quotes customer dropdown -> Assign Macro -> ClearOpenQuotesCustomerFilter" & vbNewLine & _
           "2. Right-click the Unpaid Invoices 'Apply / Clear' button -> Assign Macro -> RefreshUnpaidInvoicesList" & vbNewLine & _
           "3. Right-click the Unpaid Invoices 'Refresh' button -> Assign Macro -> RefreshUnpaidInvoicesList" & vbNewLine & _
           "4. Right-click the Unpaid 'Clear' button -> Assign Macro -> ClearUnpaidInvoicesCustomerFilter" & vbNewLine & vbNewLine & _
           "Pagination is built in -- double-click PREV / NEXT cells on each panel." & vbNewLine & _
           "Then click Refresh on both panels.", vbInformation, "Setup Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Setup error: " & Err.Description, vbCritical
End Sub

' ----- Open Quotes Customer row (row 27) -----
Private Sub SetupOpenQuotesCustomerRow(ws As Worksheet)
    On Error Resume Next
    
    ws.Range("B27").Value = "Customer:"
    ws.Range("B27").Font.bold = True
    ws.Range("B27").HorizontalAlignment = xlRight
    ws.Range("B27").VerticalAlignment = xlCenter
    
    ws.Range("C27:D27").UnMerge
    ws.Range("C27:D27").Merge
    With ws.Range("C27:D27")
        .Value = ""
        .Interior.Color = RGB(240, 245, 250)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 200, 220)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.bold = False
    End With
    
    ws.Range("E27").Value = "(leave blank for all)"
    ws.Range("E27").Font.Italic = True
    ws.Range("E27").Font.Color = RGB(120, 120, 120)
    ws.Range("E27").Font.Size = 9
    ws.Range("E27").HorizontalAlignment = xlLeft
    ws.Range("E27").VerticalAlignment = xlCenter
    
    ws.rows("27").rowHeight = 22
    
    On Error GoTo 0
End Sub

' ----- Unpaid Invoices full panel (rows 49 - 69) -----
Private Sub SetupUnpaidInvoicesPanel(ws As Worksheet)
    On Error Resume Next
    
    ' Row 49 - Filter row (From / To)
    ws.Range("B49").Value = "From:"
    ws.Range("B49").Font.bold = True
    ws.Range("B49").HorizontalAlignment = xlRight
    ws.Range("B49").VerticalAlignment = xlCenter
    
    With ws.Range("C49")
        .Value = ""
        .NumberFormat = "YYYY/MM/DD"
        .Interior.Color = RGB(240, 245, 250)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 200, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("E49").Value = "To:"
    ws.Range("E49").Font.bold = True
    ws.Range("E49").HorizontalAlignment = xlRight
    ws.Range("E49").VerticalAlignment = xlCenter
    
    With ws.Range("F49")
        .Value = ""
        .NumberFormat = "YYYY/MM/DD"
        .Interior.Color = RGB(240, 245, 250)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 200, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.rows("49").rowHeight = 22
    
    ' Row 50 - Customer row
    ws.Range("B50").Value = "Customer:"
    ws.Range("B50").Font.bold = True
    ws.Range("B50").HorizontalAlignment = xlRight
    ws.Range("B50").VerticalAlignment = xlCenter
    
    ws.Range("C50:D50").UnMerge
    ws.Range("C50:D50").Merge
    With ws.Range("C50:D50")
        .Value = ""
        .Interior.Color = RGB(240, 245, 250)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 200, 220)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("E50").Value = "(leave blank for all)"
    ws.Range("E50").Font.Italic = True
    ws.Range("E50").Font.Color = RGB(120, 120, 120)
    ws.Range("E50").Font.Size = 9
    ws.Range("E50").HorizontalAlignment = xlLeft
    ws.Range("E50").VerticalAlignment = xlCenter
    
    ws.rows("50").rowHeight = 22
    
    ' Row 51 - Blue table header
    With ws.Range("B51:H51")
        .Interior.Color = RGB(46, 135, 176)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .rowHeight = 24
    End With
    ws.Range("B51").Value = "Date"
    ws.Range("C51").Value = "Invoice #"
    ws.Range("D51").Value = "Customer"
    ws.Range("E51").Value = "Patient"
    ' F51, G51, H51 will be overwritten by the pagination labels at refresh time
    ws.Range("F51").Value = ""
    ws.Range("G51").Value = ""
    ws.Range("H51").Value = ""
    ws.rows("51").rowHeight = 24
    
    ' Row 52 - spacer
    ws.rows("52").rowHeight = 6
    
    On Error GoTo 0
End Sub

' ----- NEW: Pagination cell setup -----
' We render the actual pager labels at refresh time, but we initialise
' the state cells here so first-load logic works.
Private Sub SetupPaginationCells(ws As Worksheet)
    On Error Resume Next
    
    ' --- Open Quotes pager (row 28, F:H) ---
    ws.Range("F28:H28").ClearContents
    ws.Range("F28").Value = ChrW(9664) & " Prev"
    ws.Range("G28").Value = "Page 1 of 1"
    ws.Range("H28").Value = "Next " & ChrW(9654)
    With ws.Range("F28:H28")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.rows("28").rowHeight = 22
    
    ' Hidden state cell for Open Quotes current page
    ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Value = 1
    ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Font.Color = _
        ws.Range(MENU_OPEN_QUOTES_STATE_CELL).Interior.Color
    
    ' --- Unpaid Invoices pager labels live on row 51 (the blue header row) ---
    ' Initial values; RenderUnpaidInvoicesPagerLabel will overwrite at refresh.
    ws.Range("F51").Value = ChrW(9664) & " Prev"
    ws.Range("G51").Value = "Page 1 of 1"
    ws.Range("H51").Value = "Next " & ChrW(9654)
    ws.Range("F51").HorizontalAlignment = xlCenter
    ws.Range("G51").HorizontalAlignment = xlCenter
    ws.Range("H51").HorizontalAlignment = xlCenter
    
    ' Hidden state cell for Unpaid Invoices current page
    ws.Range(MENU_UNPAID_INV_STATE_CELL).Value = 1
    ws.Range(MENU_UNPAID_INV_STATE_CELL).Font.Color = _
        ws.Range(MENU_UNPAID_INV_STATE_CELL).Interior.Color
    
    On Error GoTo 0
End Sub

' ----- Data Validation: pull customer list from Customers!A:A -----
Private Sub SetupCustomerValidation(ws As Worksheet)
    Dim wsCust As Worksheet
    Dim lastRow As Long
    Dim listFormula As String
    
    On Error Resume Next
    
    If Not SheetExists(SHEET_CUSTOMERS) Then Exit Sub
    Set wsCust = ThisWorkbook.Sheets(SHEET_CUSTOMERS)
    lastRow = wsCust.Cells(wsCust.rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    listFormula = "=" & SHEET_CUSTOMERS & "!$A$2:$A$" & lastRow
    
    ' Open Quotes customer dropdown
    With ws.Range(MENU_OPEN_QUOTES_CUST_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
    
    ' Unpaid Invoices customer dropdown
    With ws.Range(MENU_UNPAID_INV_CUST_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
    
    On Error GoTo 0
End Sub

