Attribute VB_Name = "modInvoice"
Option Explicit

' =====================================================
' ===== GET NEXT INVOICE NUMBER =====
' =====================================================

Public Function GetNextInvoiceNumber() As String
    Dim nextInv As String
    
    nextInv = GetNextInvoiceNumberFromTracking()
    
    If nextInv <> "ERROR" Then
        GetNextInvoiceNumber = nextInv
        Exit Function
    End If
    
    Dim lastNum As Long
    
    EnsureSettingsSheet
    If Not SheetExists(SHEET_SETTINGS) Then
        ShowError "Settings sheet not found."
        GetNextInvoiceNumber = "ERROR"
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    lastNum = ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B2").Value
    GetNextInvoiceNumber = INVOICE_PREFIX & "-" & Format(lastNum + 1, "0000")
    Exit Function

ErrorHandler:
    ShowError "Error generating invoice number: " & Err.Description
    GetNextInvoiceNumber = "ERROR"
End Function

' =====================================================
' ===== CREATE NEW INVOICE =====
' =====================================================

Public Sub CreateNewInvoice()
    Dim ws As Worksheet
    Dim i As Integer
    Dim newInvoiceNum As String
    Dim invoiceDate As Date
    Dim pickedDate As Date
    
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    
    ' --- Ask for the invoice date FIRST. If user cancels, abort cleanly. ---
    pickedDate = PickDate(Date)
    If pickedDate = 0 Then
        If Not AskYesNo("No date selected. Use today's date (" & Format(Date, "YYYY/MM/DD") & ")?", _
                        "New Invoice") Then
            Exit Sub
        End If
        invoiceDate = Date
    Else
        invoiceDate = pickedDate
    End If
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Creating new invoice..."
    
    On Error Resume Next
    ws.Range(CELL_CUSTOMER).ClearContents
    ws.Range("C8").ClearContents
    ws.Range("C9").ClearContents
    ws.Range("C11").ClearContents
    ws.Range("D11").ClearContents
    ws.Range("G9").ClearContents
    ws.Range("G10").ClearContents
    ws.Range("G11").ClearContents
    ws.Range(CELL_APPLIANCE).ClearContents
    ws.Range(CELL_PATIENT).ClearContents
    
    ws.Range(CELL_DISCOUNT_PCT).ClearContents
    ws.Range(CELL_DISCOUNT_PCT).NumberFormat = "0%"
    On Error GoTo ErrorHandler
    
    For i = START_ROW To END_ROW
        On Error Resume Next
        ws.Cells(i, 1).ClearContents
        ws.Cells(i, 4).ClearContents
        On Error GoTo ErrorHandler
    Next i
    
    RestoreInvoiceFormulas
    
    newInvoiceNum = GetNextInvoiceNumber()
    If newInvoiceNum = "ERROR" Then GoTo CleanUp
    
    ws.Range(CELL_DOC_NUMBER).Value = newInvoiceNum
    ws.Range(CELL_DOC_DATE).Value = invoiceDate
    ws.Range(CELL_DOC_DATE).NumberFormat = "YYYY/MM/DD"
    
    RestoreExcelState
    
    ' --- BULLETPROOF: land on Invoice as the absolute last action ---
    ForceActivateSheet ws, CELL_CUSTOMER
    
    ShowInfo "New Invoice Created: " & newInvoiceNum & vbNewLine & _
             "Date: " & Format(invoiceDate, "YYYY/MM/DD"), "New Invoice"
    
    ' Activate AGAIN after the MsgBox returns, in case anything moved focus.
    ForceActivateSheet ws, CELL_CUSTOMER
    Exit Sub

CleanUp:
    RestoreExcelState
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error creating new invoice: " & Err.Description
End Sub

' =====================================================
' ===== CONVERT QUOTE TO INVOICE =====
' =====================================================
' Quote number is NOT displayed on the invoice sheet.
' Stored only in the InvoiceTracking file for audit trail.
'
' BUG FIX (v1.5.0):
'   The tracker open/close cycles inside UpdateQuoteStatus and
'   SaveAndVerifyInvoice can leave Excel on a different sheet
'   (e.g. the hidden Reference sheet was last touched while
'   ScreenUpdating was off). We now:
'     1. Remember the original active sheet at entry
'     2. Use Application.Goto (not .Activate) for the final landing
'     3. Re-issue the activation AFTER the success/failure popup
'
' Application.Goto is the bulletproof way: it forces the sheet
' visible, scrolls to the target range, and selects it — all in
' one call that survives EnableEvents toggling.

Public Sub ConvertToInvoice()
    Dim wsQuote As Worksheet, wsInvoice As Worksheet
    Dim invoiceNum As String, quoteNum As String
    Dim invoiceDate As Date
    Dim pickedDate As Date
    Dim vSubTotal As Variant, vDiscount As Variant, vVAT As Variant, vTotal As Variant
    Dim customer As String, patient As String
    Dim saveOk As Boolean
    
    If Not SheetExists(SHEET_QUOTE) Or Not SheetExists(SHEET_INVOICE) Then
        ShowError "Quote or Invoice sheet not found."
        Exit Sub
    End If
    
    Set wsQuote = ThisWorkbook.Sheets(SHEET_QUOTE)
    Set wsInvoice = ThisWorkbook.Sheets(SHEET_INVOICE)
    
    If Trim(wsQuote.Range(CELL_CUSTOMER).Value) = "" Then
        ShowWarning "No quote data to convert.", "No Data"
        Exit Sub
    End If
    
    quoteNum = SafeString(wsQuote.Range(CELL_DOC_NUMBER).Value)
    customer = SafeString(wsQuote.Range(CELL_CUSTOMER).Value)
    patient = SafeString(wsQuote.Range(CELL_PATIENT).Value)
    
    ' --- Ask for the invoice date. If user cancels, abort cleanly. ---
    pickedDate = PickDate(Date)
    If pickedDate = 0 Then
        If Not AskYesNo("No date selected. Use today's date (" & Format(Date, "YYYY/MM/DD") & ")?", _
                        "Convert to Invoice") Then
            Exit Sub
        End If
        invoiceDate = Date
    Else
        invoiceDate = pickedDate
    End If
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Converting to invoice..."
    
    invoiceNum = GetNextInvoiceNumber()
    If invoiceNum = "ERROR" Then GoTo CleanUp
    
    RestoreInvoiceFormulas
    
    wsInvoice.Range(CELL_CUSTOMER).Value = wsQuote.Range(CELL_CUSTOMER).Value
    wsInvoice.Range(CELL_DOC_DATE).Value = invoiceDate
    wsInvoice.Range(CELL_DOC_DATE).NumberFormat = "YYYY/MM/DD"
    wsInvoice.Range(CELL_DOC_NUMBER).Value = invoiceNum
    wsInvoice.Range(CELL_APPLIANCE).Value = wsQuote.Range(CELL_APPLIANCE).Value
    wsInvoice.Range(CELL_PATIENT).Value = wsQuote.Range(CELL_PATIENT).Value
    
    wsQuote.Range("A" & START_ROW & ":H" & END_ROW).Copy
    wsInvoice.Range("A" & START_ROW).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    wsInvoice.Range(CELL_DISCOUNT_PCT).Value = wsQuote.Range(CELL_DISCOUNT_PCT).Value
    wsInvoice.Range(CELL_DISCOUNT_PCT).NumberFormat = "0%"
    
    Application.Calculate
    
    vSubTotal = wsInvoice.Range(CELL_SUMMARY_SUBTOT).Value
    vDiscount = wsInvoice.Range(CELL_SUMMARY_DISC).Value
    vVAT = wsInvoice.Range(CELL_SUMMARY_VAT).Value
    vTotal = wsInvoice.Range(CELL_SUMMARY_TOTAL).Value
    
    ' ---- Tracker operations: these open/close external workbooks ----
    UpdateQuoteStatus quoteNum, STATUS_QUOTE_CONVERTED, invoiceNum
    
    saveOk = SaveAndVerifyInvoice(invoiceNum, quoteNum, invoiceDate, _
                                   customer, patient, _
                                   SafeString(wsQuote.Range(CELL_APPLIANCE).Value), _
                                   vSubTotal, vDiscount, vVAT, vTotal, _
                                   STATUS_INVOICE_UNPAID, Empty, "")
    
    PersistInvoiceCounter invoiceNum
    RestoreExcelState
    
    ' --- BULLETPROOF: force Invoice sheet active BEFORE the popup ---
    ForceActivateSheet wsInvoice, CELL_CUSTOMER
    
    If saveOk Then
        ShowVerifiedConfirmation "Invoice", invoiceNum, customer, patient, _
                                 r2(vTotal), _
                                 "Created from Quote " & quoteNum & vbNewLine & _
                                 "Date: " & Format(invoiceDate, "YYYY/MM/DD")
    Else
        ShowVerifyFailure "Invoice", invoiceNum, customer, patient, r2(vTotal)
    End If
    
    ' --- AND AGAIN after the popup, so we definitely end on Invoice ---
    ForceActivateSheet wsInvoice, CELL_CUSTOMER
    Exit Sub

CleanUp:
    RestoreExcelState
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error converting to invoice: " & Err.Description
End Sub

' =====================================================
' ===== FORCE ACTIVATE SHEET (bulletproof) =====
' =====================================================
' Application.Goto is more aggressive than .Activate:
'   - Makes the sheet visible if hidden (xlSheetHidden, not VeryHidden)
'   - Scrolls the target range into view
'   - Selects the target range
'   - Survives EnableEvents toggling
' Use this whenever a multi-step flow has touched external workbooks
' and we need to guarantee we end on the right sheet+cell.

Private Sub ForceActivateSheet(ws As Worksheet, ByVal targetCellAddress As String)
    On Error Resume Next
    
    ' Don't fight EnableEvents
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Application.Goto is the most reliable way to land on a sheet+cell
    Application.Goto ws.Range(targetCellAddress), Scroll:=True
    
    ' Belt + braces
    ws.Activate
    ws.Range(targetCellAddress).Select
    
    On Error GoTo 0
End Sub

' =====================================================
' ===== EXPORT INVOICE TO PDF =====
' =====================================================

Public Sub ExportInvoiceToPDF()
    Dim ws As Worksheet
    Dim savePath As String, fileName As String
    Dim patientName As String, invoiceNum As String
    
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    
    If Trim(ws.Range(CELL_DOC_NUMBER).Value) = "" Then
        ShowWarning "Please create an invoice first.", "Missing Invoice"
        Exit Sub
    End If
    
    patientName = CleanFileName(SafeString(ws.Range(CELL_PATIENT).Value))
    If patientName = "" Then patientName = CleanFileName(SafeString(ws.Range(CELL_CUSTOMER).Value))
    If patientName = "" Then patientName = "Invoice"
    invoiceNum = CleanFileName(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    fileName = patientName & " - " & invoiceNum & ".pdf"
    
    savePath = GetSavePath(fileName, "PDF Files (*.pdf), *.pdf", "Save Invoice as PDF")
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    SafeStatusBar "Exporting Invoice PDF..."
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=savePath, Quality:=xlQualityStandard
    
    SafeStatusBar False
    ShowInfo "Invoice PDF exported to:" & vbNewLine & vbNewLine & savePath, "Export Complete"
    
    ' Land back on the Invoice sheet in case the export shifted focus
    ForceActivateSheet ws, CELL_CUSTOMER
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error exporting Invoice PDF: " & Err.Description
End Sub

' =====================================================
' ===== EXPORT INVOICE TO EXCEL (FULL WORKBOOK COPY) =====
' =====================================================
' Quote-number link is preserved by re-reading the existing tracker row
' (so we don't lose it when re-exporting).

Public Sub ExportInvoiceToExcel()
    Dim ws As Worksheet
    Dim savePath As String, fileName As String
    Dim patientName As String, invoiceNum As String, quoteNum As String
    Dim customer As String, patient As String
    Dim vSubTotal As Variant, vDiscount As Variant, vVAT As Variant, vTotal As Variant
    Dim invoiceDate As Date, currentDate As Date
    Dim pickedDate As Date
    Dim saveOk As Boolean
    
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    
    If Trim(ws.Range(CELL_DOC_NUMBER).Value) = "" Then
        ShowWarning "Please create an invoice first.", "Missing Invoice"
        Exit Sub
    End If
    
    If IsDate(ws.Range(CELL_DOC_DATE).Value) Then
        currentDate = CDate(ws.Range(CELL_DOC_DATE).Value)
    Else
        currentDate = Date
    End If
    
    If AskYesNo("Do you want to change the invoice date before exporting?" & vbNewLine & _
                "Current date: " & Format(currentDate, "YYYY/MM/DD")) Then
        pickedDate = PickDate(currentDate)
        ' If user cancelled, keep the existing date instead of silently picking today.
        If pickedDate > 0 Then
            invoiceDate = pickedDate
            ws.Range(CELL_DOC_DATE).Value = invoiceDate
            ws.Range(CELL_DOC_DATE).NumberFormat = "YYYY/MM/DD"
        Else
            invoiceDate = currentDate
        End If
    Else
        invoiceDate = currentDate
    End If
    
    patientName = CleanFileName(SafeString(ws.Range(CELL_PATIENT).Value))
    If patientName = "" Then patientName = CleanFileName(SafeString(ws.Range(CELL_CUSTOMER).Value))
    If patientName = "" Then patientName = "Invoice"
    invoiceNum = CleanFileName(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    fileName = patientName & " - " & invoiceNum & ".xlsm"
    
    savePath = GetSavePath(fileName, "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", "Save Invoice as Excel")
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Exporting invoice (full workbook)..."
    
    customer = SafeString(ws.Range(CELL_CUSTOMER).Value)
    patient = SafeString(ws.Range(CELL_PATIENT).Value)
    
    vSubTotal = ws.Range(CELL_SUMMARY_SUBTOT).Value
    vDiscount = ws.Range(CELL_SUMMARY_DISC).Value
    vVAT = ws.Range(CELL_SUMMARY_VAT).Value
    vTotal = ws.Range(CELL_SUMMARY_TOTAL).Value
    
    quoteNum = LookupExistingInvoiceQuoteNum(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs savePath
    Application.DisplayAlerts = True
    
    ' ---- Save AND verify ----
    saveOk = SaveAndVerifyInvoice(SafeString(ws.Range(CELL_DOC_NUMBER).Value), quoteNum, _
                                   SafeDate(ws.Range(CELL_DOC_DATE).Value), _
                                   customer, patient, _
                                   SafeString(ws.Range(CELL_APPLIANCE).Value), _
                                   vSubTotal, vDiscount, vVAT, vTotal, _
                                   STATUS_INVOICE_UNPAID, Empty, savePath)
    
    PersistInvoiceCounter SafeString(ws.Range(CELL_DOC_NUMBER).Value)
    RestoreExcelState
    
    ' Force land on Invoice before AND after the popup
    ForceActivateSheet ws, CELL_CUSTOMER
    
    If saveOk Then
        ShowVerifiedConfirmation "Invoice", invoiceNum, customer, patient, _
                                 r2(vTotal), _
                                 "Exported to: " & savePath & vbNewLine & _
                                 "Date synced: " & Format(SafeDate(ws.Range(CELL_DOC_DATE).Value), "YYYY/MM/DD")
    Else
        ShowVerifyFailure "Invoice", invoiceNum, customer, patient, r2(vTotal)
    End If
    
    ForceActivateSheet ws, CELL_CUSTOMER
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error exporting invoice: " & Err.Description
End Sub

' =====================================================
' ===== PRINT INVOICE =====
' =====================================================

Public Sub PrintInvoice()
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    ThisWorkbook.Sheets(SHEET_INVOICE).PrintOut Copies:=1, Preview:=True
    Exit Sub

ErrorHandler:
    ShowError "Error printing invoice: " & Err.Description
End Sub

