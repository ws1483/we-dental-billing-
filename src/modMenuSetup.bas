Attribute VB_Name = "modMenuSetup"
Option Explicit

' =====================================================
' ===== ONE-TIME SETUP FOR MENU SHEET PANELS =====
' Run SetupMenuPanels ONCE (Alt+F8) to:
'   - add Customer dropdown row to Open Quotes panel
'   - draw the Unpaid Invoices panel below it
'   - set up Data Validation lists from Customers sheet
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
    
    Application.ScreenUpdating = True
    MsgBox "Menu panels set up successfully." & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Right-click the small 'Clear' button next to the Open Quotes customer dropdown -> Assign Macro -> ClearOpenQuotesCustomerFilter" & vbNewLine & _
           "2. Right-click the Unpaid Invoices 'Apply / Clear' button -> Assign Macro -> RefreshUnpaidInvoicesList" & vbNewLine & _
           "3. Right-click the Unpaid Invoices 'Refresh' button -> Assign Macro -> RefreshUnpaidInvoicesList" & vbNewLine & _
           "4. Right-click the small Unpaid 'Clear' button -> Assign Macro -> ClearUnpaidInvoicesCustomerFilter" & vbNewLine & vbNewLine & _
           "Then click Refresh on both panels.", vbInformation, "Setup Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Setup error: " & Err.Description, vbCritical
End Sub

' ----- Open Quotes Customer row (row 27) -----
Private Sub SetupOpenQuotesCustomerRow(ws As Worksheet)
    On Error Resume Next
    
    ' Row 27 - Customer filter row
    ws.Range("B27").Value = "Customer:"
    ws.Range("B27").Font.bold = True
    ws.Range("B27").HorizontalAlignment = xlRight
    ws.Range("B27").VerticalAlignment = xlCenter
    
    ' Dropdown cell C27:D27
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
    
    ' Label hint
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
    
    ' Row 50 - Customer row + title
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
    
    ' Row 50 also carries the title (column F onwards or above the header)
    ' We'll use row 50 column F onward for the title text via UpdateUnpaidInvoicesHeaderRange
    ' (it writes into row MENU_UNPAID_INV_HEADER_ROW - 1, i.e. row 50, B:F)
    ' To avoid overlap with Customer label, we keep title in F:H during refresh.
    
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
    ws.Range("F51").Value = "Appliance"
    ws.Range("G51").Value = "Balance"
    ws.Range("G51").HorizontalAlignment = xlRight
    ws.Range("H51").Value = "Age"
    ws.Range("H51").HorizontalAlignment = xlCenter
    ws.rows("51").rowHeight = 24
    
    ' Row 52 - spacer (matches Open Quotes layout)
    ws.rows("52").rowHeight = 6
    
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

