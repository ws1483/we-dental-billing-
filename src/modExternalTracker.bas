Attribute VB_Name = "modExternalTracker"
Option Explicit

' =====================================================
' ===== TRACKER RECORD TYPES =====
' =====================================================

Public Type InvoiceRecord
    found As Boolean
    status As String
    paidDate As Variant
    filePath As String
    totalIncl As Double
    paidAmount As Double
    creditAmount As Double
End Type

Public Type QuoteRecord
    found As Boolean
    filePath As String
End Type

' =====================================================
' ===== ONE-PER-SESSION WARNING FLAGS =====
' =====================================================
' We only nag the user ONCE per Excel session if the tracker
' has higher document numbers than Settings says it should.
' Resetting these flags requires reopening the workbook.

Private mWarnedQuoteCounter   As Boolean
Private mWarnedInvoiceCounter As Boolean

' =====================================================
' ===== UPDATE QUOTE TO TRACKER (from current sheet) =====
' =====================================================

Public Sub UpdateQuoteToTracker()
    Dim ws As Worksheet
    Dim quoteNum As String
    Dim vSubTotal As Double, vDiscount As Double, vVAT As Double, vTotal As Double
    Dim qr As QuoteRecord
    
    If Not SheetExists(SHEET_QUOTE) Then
        ShowError "Quote sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_QUOTE)
    quoteNum = Trim(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    
    If quoteNum = "" Then
        ShowWarning "No quote number found on the Quote sheet." & vbNewLine & _
                    "Please create or load a quote first.", "No Quote"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Updating quote to tracker..."
    
    vSubTotal = r2(ws.Range(CELL_SUMMARY_SUBTOT).Value)
    vDiscount = r2(ws.Range(CELL_SUMMARY_DISC).Value)
    vVAT = r2(ws.Range(CELL_SUMMARY_VAT).Value)
    vTotal = r2(ws.Range(CELL_SUMMARY_TOTAL).Value)
    
    qr = ReadQuoteRecord(quoteNum)
    
    TrackQuote quoteNum, _
               SafeDate(ws.Range(CELL_DOC_DATE).Value), _
               SafeString(ws.Range(CELL_CUSTOMER).Value), _
               SafeString(ws.Range(CELL_PATIENT).Value), _
               SafeString(ws.Range(CELL_APPLIANCE).Value), _
               vSubTotal, vDiscount, vVAT, vTotal, _
               STATUS_QUOTE_OPEN, "", qr.filePath
    
    RestoreExcelState
    
    ShowInfo "Quote " & quoteNum & " updated in tracker." & vbNewLine & vbNewLine & _
             "Customer: " & SafeString(ws.Range(CELL_CUSTOMER).Value) & vbNewLine & _
             "Patient: " & SafeString(ws.Range(CELL_PATIENT).Value) & vbNewLine & _
             "Total: " & FmtR(vTotal), _
             "Tracker Updated"
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error updating quote to tracker: " & Err.Description
End Sub

' =====================================================
' ===== UPDATE INVOICE TO TRACKER (from current sheet) =====
' =====================================================

Public Sub UpdateInvoiceToTracker()
    Dim ws As Worksheet
    Dim invoiceNum As String, quoteNum As String
    Dim vSubTotal As Double, vDiscount As Double, vVAT As Double, vTotal As Double
    Dim ir As InvoiceRecord
    
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    invoiceNum = Trim(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    
    If invoiceNum = "" Then
        ShowWarning "No invoice number found on the Invoice sheet." & vbNewLine & _
                    "Please create or load an invoice first.", "No Invoice"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Updating invoice to tracker..."
    
    vSubTotal = r2(ws.Range(CELL_SUMMARY_SUBTOT).Value)
    vDiscount = r2(ws.Range(CELL_SUMMARY_DISC).Value)
    vVAT = r2(ws.Range(CELL_SUMMARY_VAT).Value)
    vTotal = r2(ws.Range(CELL_SUMMARY_TOTAL).Value)
    
    ' Quote number is no longer displayed on the invoice sheet.
    ' Preserve whatever quote link is already in the tracker.
    quoteNum = LookupExistingInvoiceQuoteNum(invoiceNum)
    
    ir = ReadInvoiceRecord(invoiceNum)
    If ir.status = "" Then ir.status = STATUS_INVOICE_UNPAID
    
    TrackInvoice invoiceNum, quoteNum, _
                 SafeDate(ws.Range(CELL_DOC_DATE).Value), _
                 SafeString(ws.Range(CELL_CUSTOMER).Value), _
                 SafeString(ws.Range(CELL_PATIENT).Value), _
                 SafeString(ws.Range(CELL_APPLIANCE).Value), _
                 vSubTotal, vDiscount, vVAT, vTotal, _
                 ir.status, ir.paidDate, ir.filePath
    
    RestoreExcelState
    
    ShowInfo "Invoice " & invoiceNum & " updated in tracker." & vbNewLine & vbNewLine & _
             "Customer: " & SafeString(ws.Range(CELL_CUSTOMER).Value) & vbNewLine & _
             "Patient: " & SafeString(ws.Range(CELL_PATIENT).Value) & vbNewLine & _
             "Date: " & Format(SafeDate(ws.Range(CELL_DOC_DATE).Value), "YYYY/MM/DD") & vbNewLine & _
             "Total: " & FmtR(vTotal) & vbNewLine & _
             "Status: " & ir.status, _
             "Tracker Updated"
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error updating invoice to tracker: " & Err.Description
End Sub

' =====================================================
' ===== READ EXISTING RECORDS (SINGLE OPEN) =====
' =====================================================

Private Function ReadQuoteRecord(ByVal quoteNum As String) As QuoteRecord
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    Dim rec As QuoteRecord
    
    rec.found = False
    rec.filePath = ""
    
    On Error GoTo SafeExit
    
    trackingPath = GetQuoteTrackingPath()
    If trackingPath = "" Then GoTo SafeExit
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then GoTo SafeExit
    
    If SheetExistsInWorkbook(TRACK_SHEET_QUOTES, wbTrack) Then
        Set wsTrack = wbTrack.Sheets(TRACK_SHEET_QUOTES)
        targetRow = FindDocumentRow(wsTrack, quoteNum, COL_Q_NUMBER)
        If targetRow > 0 Then
            rec.found = True
            rec.filePath = SafeString(wsTrack.Cells(targetRow, COL_Q_FILEPATH).Value)
        End If
    End If
    
SafeExit:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ReadQuoteRecord = rec
End Function

Public Function ReadInvoiceRecord(ByVal invoiceNum As String) As InvoiceRecord
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    Dim rec As InvoiceRecord
    
    rec.found = False
    rec.status = ""
    rec.paidDate = Empty
    rec.filePath = ""
    rec.totalIncl = 0
    rec.paidAmount = 0
    rec.creditAmount = 0
    
    On Error GoTo SafeExit
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then GoTo SafeExit
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then GoTo SafeExit
    
    If SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
        targetRow = FindDocumentRow(wsTrack, invoiceNum, COL_I_NUMBER)
        If targetRow > 0 Then
            rec.found = True
            rec.status = Trim(SafeString(wsTrack.Cells(targetRow, COL_I_STATUS).Value))
            If IsDate(wsTrack.Cells(targetRow, COL_I_PAIDDATE).Value) Then
                rec.paidDate = wsTrack.Cells(targetRow, COL_I_PAIDDATE).Value
            End If
            rec.filePath = SafeString(wsTrack.Cells(targetRow, COL_I_FILEPATH).Value)
            rec.totalIncl = r2(wsTrack.Cells(targetRow, COL_I_TOTAL).Value)
            rec.paidAmount = r2(wsTrack.Cells(targetRow, COL_I_PAIDAMOUNT).Value)
            rec.creditAmount = r2(wsTrack.Cells(targetRow, COL_I_CREDITAMT).Value)
        End If
    End If
    
SafeExit:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ReadInvoiceRecord = rec
End Function

' =====================================================
' ===== APPLY PARTIAL PAYMENT =====
' =====================================================
' Records a payment, recalculates status, and AUDIT-LOGS any overpayment
' so the user can see they were capped (instead of silently losing money).

Public Sub ApplyPartialPayment(ByVal invoiceNum As String, _
                               ByVal paymentAmount As Double, _
                               Optional ByVal paymentDate As Variant)
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    Dim total As Double, credit As Double
    Dim existingPaid As Double, newPaid As Double
    Dim due As Double, newStatus As String
    Dim overpayment As Double
    Dim customer As String, patient As String
    Dim usePaymentDate As Date
    
    If Trim(invoiceNum) = "" Then
        ShowError "Invoice number is empty."
        Exit Sub
    End If
    
    If paymentAmount <= 0 Then
        ShowError "Payment amount must be greater than zero."
        Exit Sub
    End If
    
    If IsMissing(paymentDate) Then
        usePaymentDate = Date
    ElseIf IsDate(paymentDate) Then
        usePaymentDate = CDate(paymentDate)
    Else
        usePaymentDate = Date
    End If
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath)
    If wbTrack Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        ShowError "Invoice tracking sheet not found."
        Exit Sub
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    EnsurePaidAmountHeader wsTrack
    
    targetRow = FindDocumentRow(wsTrack, invoiceNum, COL_I_NUMBER)
    
    If targetRow = 0 Then
        CloseTrackingWorkbook wbTrack, False
        ShowError "Invoice " & invoiceNum & " not found in tracking."
        Exit Sub
    End If
    
    total = r2(wsTrack.Cells(targetRow, COL_I_TOTAL).Value)
    credit = r2(wsTrack.Cells(targetRow, COL_I_CREDITAMT).Value)
    existingPaid = r2(wsTrack.Cells(targetRow, COL_I_PAIDAMOUNT).Value)
    customer = SafeString(wsTrack.Cells(targetRow, COL_I_CUSTOMER).Value)
    patient = SafeString(wsTrack.Cells(targetRow, COL_I_PATIENT).Value)
    
    newPaid = Round(existingPaid + r2(paymentAmount), 2)
    
    due = Round(total - credit, 2)
    If due < 0 Then due = 0
    
    overpayment = 0
    
    If newPaid <= 0 Then
        newStatus = STATUS_INVOICE_UNPAID
    ElseIf newPaid >= due Then
        newStatus = STATUS_INVOICE_PAID
        If newPaid > due Then
            ' --- Cap the recorded paid amount at the due total, but
            '     audit-log the excess so it's visible to anyone reviewing
            '     the log later. The user-facing message also surfaces it.
            overpayment = Round(newPaid - due, 2)
            newPaid = due
            On Error Resume Next
            LogAudit "ApplyPartialPayment", invoiceNum, customer, patient, total, _
                     "OVERPAYMENT", _
                     "Payment of " & FmtR(paymentAmount) & " exceeded balance by " & FmtR(overpayment) & _
                     " — recorded amount capped at " & FmtR(due)
            On Error GoTo ErrorHandler
        End If
    Else
        newStatus = STATUS_INVOICE_PARTIAL
    End If
    
    With wsTrack
        .Cells(targetRow, COL_I_PAIDAMOUNT).Value = newPaid
        .Cells(targetRow, COL_I_PAIDAMOUNT).NumberFormat = "R#,##0.00"
        
        .Cells(targetRow, COL_I_STATUS).Value = newStatus
        
        If newStatus = STATUS_INVOICE_PAID Then
            .Cells(targetRow, COL_I_PAIDDATE).Value = usePaymentDate
            .Cells(targetRow, COL_I_PAIDDATE).NumberFormat = "YYYY/MM/DD"
        End If
        
        .Cells(targetRow, COL_I_LASTMOD).Value = Now
        .Cells(targetRow, COL_I_LASTMOD).NumberFormat = "YYYY/MM/DD HH:MM"
    End With
    
    CloseTrackingWorkbook wbTrack, True
    
    Dim remaining As Double
    remaining = Round(due - newPaid, 2)
    If remaining < 0 Then remaining = 0
    
    Dim msg As String
    msg = "Payment recorded for " & invoiceNum & "." & vbNewLine & vbNewLine & _
          "Payment:    " & FmtR(paymentAmount) & vbNewLine & _
          "Paid total: " & FmtR(newPaid) & vbNewLine & _
          "Remaining:  " & FmtR(remaining) & vbNewLine & _
          "Status:     " & newStatus
    
    If overpayment > 0 Then
        msg = msg & vbNewLine & vbNewLine & _
              "Note: payment exceeded balance by " & FmtR(overpayment) & "." & vbNewLine & _
              "The recorded paid amount has been capped at the invoice total." & vbNewLine & _
              "See AuditLog sheet (hidden) for details."
    End If
    
    ShowInfo msg, "Payment Applied"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ShowError "Error applying payment: " & Err.Description
End Sub

Private Sub EnsurePaidAmountHeader(wsTrack As Worksheet)
    On Error Resume Next
    If Trim(CStr(wsTrack.Cells(1, COL_I_PAIDAMOUNT).Value)) = "" Then
        wsTrack.Cells(1, COL_I_PAIDAMOUNT).Value = "PaidAmount"
        wsTrack.Cells(1, COL_I_PAIDAMOUNT).Font.bold = True
        wsTrack.Cells(1, COL_I_PAIDAMOUNT).Interior.Color = RGB(0, 176, 80)
        wsTrack.Cells(1, COL_I_PAIDAMOUNT).Font.Color = RGB(255, 255, 255)
        wsTrack.Columns(COL_I_PAIDAMOUNT).AutoFit
    End If
    On Error GoTo 0
End Sub

' =====================================================
' ===== GET TRACKING PATHS (WITH MANUAL BROWSE) =====
' =====================================================

Public Function GetQuoteTrackingPath() As String
    Dim savedPath As String
    
    EnsureSettingsSheet
    
    On Error Resume Next
    savedPath = Trim(ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B3").Value)
    On Error GoTo 0
    
    If savedPath <> "" Then
        If FileExists(savedPath) Then
            GetQuoteTrackingPath = savedPath
            Exit Function
        End If
    End If
    
    savedPath = PromptForTrackingFile( _
        "Quote Tracking file not found." & vbNewLine & vbNewLine & _
        "Please locate your Quote Tracking file.", _
        "Locate Quote Tracking File")
    
    If savedPath <> "" Then SaveTrackingPath "B3", savedPath
    
    GetQuoteTrackingPath = savedPath
End Function

Public Function GetInvoiceTrackingPath() As String
    Dim savedPath As String
    
    EnsureSettingsSheet
    
    On Error Resume Next
    savedPath = Trim(ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B4").Value)
    On Error GoTo 0
    
    If savedPath <> "" Then
        If FileExists(savedPath) Then
            GetInvoiceTrackingPath = savedPath
            Exit Function
        End If
    End If
    
    savedPath = PromptForTrackingFile( _
        "Invoice Tracking file not found." & vbNewLine & vbNewLine & _
        "Please locate your Invoice Tracking file.", _
        "Locate Invoice Tracking File")
    
    If savedPath <> "" Then SaveTrackingPath "B4", savedPath
    
    GetInvoiceTrackingPath = savedPath
End Function

Private Function PromptForTrackingFile(promptMessage As String, dialogTitle As String) As String
    Dim filePath As String
    
    MsgBox promptMessage, vbExclamation, dialogTitle
    filePath = GetOpenPath("Excel Files (*.xls*), *.xls*", dialogTitle)
    
    If filePath = "" Then
        PromptForTrackingFile = ""
    Else
        PromptForTrackingFile = filePath
    End If
End Function

Private Sub SaveTrackingPath(cell As String, filePath As String)
    EnsureSettingsSheet
    On Error Resume Next
    If SheetExists(SHEET_SETTINGS) Then
        ThisWorkbook.Sheets(SHEET_SETTINGS).Range(cell).Value = filePath
    End If
    On Error GoTo 0
End Sub

Public Sub ResetTrackingPaths()
    EnsureSettingsSheet
    On Error Resume Next
    If SheetExists(SHEET_SETTINGS) Then
        ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B3").ClearContents
        ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B4").ClearContents
    End If
    On Error GoTo 0
    
    ShowInfo "Tracking file paths have been reset." & vbNewLine & _
             "You will be prompted to locate the files next time.", _
             "Paths Reset"
End Sub

Public Sub ChangeQuoteTrackingPath()
    Dim filePath As String
    filePath = GetOpenPath("Excel Files (*.xls*), *.xls*", "Select Quote Tracking File")
    
    If filePath <> "" Then
        SaveTrackingPath "B3", filePath
        ShowInfo "Quote tracking path updated to:" & vbNewLine & vbNewLine & filePath, "Path Updated"
    End If
End Sub

Public Sub ChangeInvoiceTrackingPath()
    Dim filePath As String
    filePath = GetOpenPath("Excel Files (*.xls*), *.xls*", "Select Invoice Tracking File")
    
    If filePath <> "" Then
        SaveTrackingPath "B4", filePath
        ShowInfo "Invoice tracking path updated to:" & vbNewLine & vbNewLine & filePath, "Path Updated"
    End If
End Sub

' =====================================================
' ===== CREATE TRACKING FILES =====
' =====================================================

Public Sub CreateQuoteTrackingFile()
    Dim savePath As String, newWb As Workbook, ws As Worksheet
    
    savePath = GetSavePath("QuoteTracking.xlsx", _
               "Excel Files (*.xlsx), *.xlsx", _
               "Choose Where to Save Quote Tracking File")
    
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set newWb = Workbooks.Add
    Set ws = newWb.Sheets(1)
    ws.Name = TRACK_SHEET_QUOTES
    
    CreateQuoteTrackingHeaders ws
    
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close saveChanges:=False
    Application.DisplayAlerts = True
    
    SaveTrackingPath "B3", savePath
    
    ShowInfo "Quote Tracking file created at:" & vbNewLine & vbNewLine & savePath, "File Created"
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error creating quote tracking file: " & Err.Description
End Sub

Public Sub CreateInvoiceTrackingFile()
    Dim savePath As String, newWb As Workbook, ws As Worksheet
    
    savePath = GetSavePath("InvoiceTracking.xlsx", _
               "Excel Files (*.xlsx), *.xlsx", _
               "Choose Where to Save Invoice Tracking File")
    
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set newWb = Workbooks.Add
    Set ws = newWb.Sheets(1)
    ws.Name = TRACK_SHEET_INVOICES
    
    CreateInvoiceTrackingHeaders ws
    
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close saveChanges:=False
    Application.DisplayAlerts = True
    
    SaveTrackingPath "B4", savePath
    
    ShowInfo "Invoice Tracking file created at:" & vbNewLine & vbNewLine & savePath, "File Created"
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error creating invoice tracking file: " & Err.Description
End Sub

' =====================================================
' ===== EXTERNAL WORKBOOK MANAGEMENT =====
' =====================================================
' Race-safe open: if the file is already open read-only and we need r/w,
' we ATTEMPT to close it. If closing fails (unsaved changes, lock, etc.)
' we return the existing read-only handle instead of crashing into a
' half-closed Workbooks.Open path.

Public Function OpenTrackingWorkbook(filePath As String, Optional readOnly As Boolean = False) As Workbook
    Dim wb As Workbook
    Dim existing As Workbook
    
    If filePath = "" Then
        ShowError "Tracking file path is empty."
        Set OpenTrackingWorkbook = Nothing
        Exit Function
    End If
    
    If Not FileExists(filePath) Then
        ShowError "Tracking file not found:" & vbNewLine & vbNewLine & _
                  filePath & vbNewLine & vbNewLine & _
                  "Please use 'Change Tracking Path' to update the location."
        Set OpenTrackingWorkbook = Nothing
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' --- Check if it's already open in this Excel instance ---
    For Each existing In Workbooks
        If LCase$(existing.FullName) = LCase$(filePath) Then
            ' Already open. Decide whether to reuse or re-open.
            If readOnly Or Not existing.readOnly Then
                ' Caller is happy with current openness state.
                Set OpenTrackingWorkbook = existing
                Exit Function
            End If
            
            ' Caller wants read-write but file is currently read-only.
            ' Attempt to close so we can re-open r/w.
            On Error Resume Next
            Err.Clear
            Application.ScreenUpdating = False
            existing.Close saveChanges:=False
            Application.ScreenUpdating = True
            
            If Err.Number <> 0 Then
                ' Couldn't close (locked, dirty, denied). Fall back to
                ' returning the read-only handle so the caller can at
                ' least READ. Writes will fail visibly downstream.
                LogAudit "OpenTracker", "", "", "", 0, "FAILED-REOPEN", _
                         "Could not close r/o handle to re-open r/w: " & _
                         "Err#" & Err.Number & " — " & Err.Description & _
                         " — falling back to read-only handle"
                Err.Clear
                On Error GoTo ErrorHandler
                Set OpenTrackingWorkbook = existing
                Exit Function
            End If
            On Error GoTo ErrorHandler
            
            Exit For   ' fall through to fresh Workbooks.Open below
        End If
    Next existing
    
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, UpdateLinks:=False, readOnly:=readOnly)
    Application.ScreenUpdating = True
    
    Set OpenTrackingWorkbook = wb
    Exit Function

ErrorHandler:
    Application.ScreenUpdating = True
    ShowError "Error opening tracking workbook: " & Err.Description
    Set OpenTrackingWorkbook = Nothing
End Function

Public Sub CloseTrackingWorkbook(wb As Workbook, Optional saveChanges As Boolean = True)
    If wb Is Nothing Then Exit Sub
    
    If saveChanges Then
        On Error Resume Next
        Err.Clear
        wb.Save
        If Err.Number <> 0 Then
            ' Log but don't popup — caller's verify pass will detect the failure
            ' and show a proper error to the user.
            On Error Resume Next
            LogAudit "SaveTracker", wb.Name, "", "", 0, "FAILED-SAVE", _
                     "Err#" & Err.Number & ": " & Err.Description
            On Error GoTo 0
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    On Error Resume Next
    ' Close WITHOUT save-on-close — we already saved (or tried) above.
    wb.Close saveChanges:=False
    On Error GoTo 0
End Sub

Public Function FindDocumentRow(ws As Worksheet, docNumber As String, searchColumn As Integer) As Long
    Dim foundCell As Range
    
    If Trim(docNumber) = "" Then
        FindDocumentRow = 0
        Exit Function
    End If
    
    On Error Resume Next
    Set foundCell = ws.Columns(searchColumn).Find( _
        What:=docNumber, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)
    On Error GoTo 0
    
    If foundCell Is Nothing Then
        FindDocumentRow = 0
    Else
        FindDocumentRow = foundCell.row
    End If
End Function

' =====================================================
' ===== TRACKING HEADERS =====
' =====================================================

Private Sub CreateQuoteTrackingHeaders(ws As Worksheet)
    With ws
        .Cells(1, COL_Q_NUMBER).Value = "QuoteNumber"
        .Cells(1, COL_Q_DATE).Value = "QuoteDate"
        .Cells(1, COL_Q_CUSTOMER).Value = "CustomerName"
        .Cells(1, COL_Q_PATIENT).Value = "PatientName"
        .Cells(1, COL_Q_APPLIANCE).Value = "ApplianceType"
        .Cells(1, COL_Q_SUBTOTAL).Value = "SubTotal"
        .Cells(1, COL_Q_DISCOUNT).Value = "Discount"
        .Cells(1, COL_Q_VAT).Value = "VAT"
        .Cells(1, COL_Q_TOTAL).Value = "TotalIncl"
        .Cells(1, COL_Q_STATUS).Value = "Status"
        .Cells(1, COL_Q_INVOICENUM).Value = "InvoiceNumber"
        .Cells(1, COL_Q_FILEPATH).Value = "FilePath"
        .Cells(1, COL_Q_LASTMOD).Value = "LastModified"
        
        With .Range("A1:M1")
            .Font.bold = True
            .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        .Columns("A:M").AutoFit
    End With
End Sub

Private Sub CreateInvoiceTrackingHeaders(ws As Worksheet)
    With ws
        .Cells(1, COL_I_NUMBER).Value = "InvoiceNumber"
        .Cells(1, COL_I_QUOTENUM).Value = "QuoteNumber"
        .Cells(1, COL_I_DATE).Value = "InvoiceDate"
        .Cells(1, COL_I_CUSTOMER).Value = "CustomerName"
        .Cells(1, COL_I_PATIENT).Value = "PatientName"
        .Cells(1, COL_I_APPLIANCE).Value = "ApplianceType"
        .Cells(1, COL_I_SUBTOTAL).Value = "SubTotal"
        .Cells(1, COL_I_DISCOUNT).Value = "Discount"
        .Cells(1, COL_I_VAT).Value = "VAT"
        .Cells(1, COL_I_TOTAL).Value = "TotalIncl"
        .Cells(1, COL_I_STATUS).Value = "Status"
        .Cells(1, COL_I_PAIDDATE).Value = "PaidDate"
        .Cells(1, COL_I_FILEPATH).Value = "FilePath"
        .Cells(1, COL_I_LASTMOD).Value = "LastModified"
        .Cells(1, COL_I_CREDITAMT).Value = "CreditAmount"
        .Cells(1, COL_I_CREDITREASON).Value = "CreditReason"
        .Cells(1, COL_I_PAIDAMOUNT).Value = "PaidAmount"
        
        With .Range("A1:Q1")
            .Font.bold = True
            .Interior.Color = RGB(0, 176, 80)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        .Columns("A:Q").AutoFit
    End With
End Sub

' =====================================================
' ===== EXTRACT SEQUENCE NUMBERS =====
' =====================================================
' Only returns a sequence if the row prefix matches expectedPrefix.
' Pass expectedPrefix = "" to accept any prefix.

Private Function ExtractSequence(textValue As String, Optional ByVal expectedPrefix As String = "") As Long
    Dim s As String, dashPos As Long, foundPrefix As String
    
    On Error GoTo Failed
    
    s = Trim$(textValue)
    If s = "" Then Exit Function
    
    dashPos = InStr(1, s, "-")
    If dashPos = 0 Then Exit Function
    
    foundPrefix = Left$(s, dashPos - 1)
    
    If Len(expectedPrefix) > 0 Then
        If StrComp(foundPrefix, expectedPrefix, vbTextCompare) <> 0 Then
            Exit Function
        End If
    End If
    
    s = Mid$(s, dashPos + 1)
    If s = "" Or Not IsNumeric(s) Then Exit Function
    
    ExtractSequence = CLng(s)
    Exit Function

Failed:
    ExtractSequence = 0
End Function

' =====================================================
' ===== GET NEXT QUOTE NUMBER FROM TRACKING =====
' =====================================================
' Now warns ONCE per session if the tracker is ahead of the
' Settings counter (previously popped up on every quote).

Public Function GetNextQuoteNumberFromTracking() As String
    Dim trackingPath As String, wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim quoteNum As Long, trackerMax As Long, settingsNum As Long
    Dim useSeed As Long
    
    EnsureSettingsSheet
    
    settingsNum = 0
    On Error Resume Next
    If SheetExists(SHEET_SETTINGS) Then
        settingsNum = CLng(ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B1").Value)
    End If
    On Error GoTo 0
    
    trackerMax = 0
    
    trackingPath = GetQuoteTrackingPath()
    If trackingPath = "" Then GoTo BuildNumber
    
    Set wb = OpenTrackingWorkbook(trackingPath, True)
    If wb Is Nothing Then GoTo BuildNumber
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_QUOTES, wb) Then
        CloseTrackingWorkbook wb, False
        GoTo BuildNumber
    End If
    
    Set ws = wb.Sheets(TRACK_SHEET_QUOTES)
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        quoteNum = ExtractSequence(Trim(CStr(ws.Cells(i, COL_Q_NUMBER).Value)), QUOTE_PREFIX)
        If quoteNum > trackerMax Then trackerMax = quoteNum
    Next i
    
    CloseTrackingWorkbook wb, False

BuildNumber:
    If trackerMax > settingsNum Then
        ' --- Only nag the first time per session ---
        If Not mWarnedQuoteCounter Then
            mWarnedQuoteCounter = True
            Dim msg As String
            msg = "Settings!B1 says next " & QUOTE_PREFIX & " quote is " & (settingsNum + 1) & "," & vbNewLine & _
                  "but the tracker already contains " & QUOTE_PREFIX & "-" & trackerMax & "." & vbNewLine & vbNewLine & _
                  "To avoid duplicate quote numbers, the next quote will be " & QUOTE_PREFIX & "-" & (trackerMax + 1) & "." & vbNewLine & vbNewLine & _
                  "(This warning will not repeat in this session.)" & vbNewLine & vbNewLine & _
                  "If you want to RESET numbering, set Settings!B1 to your desired starting number and " & _
                  "remove higher-numbered " & QUOTE_PREFIX & "-### rows from the tracker."
            ShowWarning msg, "Quote Number Auto-Correction"
        End If
        useSeed = trackerMax
    Else
        useSeed = settingsNum
    End If
    
    GetNextQuoteNumberFromTracking = QUOTE_PREFIX & "-" & (useSeed + 1)
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then CloseTrackingWorkbook wb, False
    
    If settingsNum > 0 Then
        GetNextQuoteNumberFromTracking = QUOTE_PREFIX & "-" & (settingsNum + 1)
    Else
        GetNextQuoteNumberFromTracking = "ERROR"
    End If
End Function

' =====================================================
' ===== GET NEXT INVOICE NUMBER FROM TRACKING =====
' =====================================================
' Same warn-once behaviour as the quote counter.

Public Function GetNextInvoiceNumberFromTracking() As String
    Dim trackingPath As String, wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim invNum As Long, trackerMax As Long, settingsNum As Long
    Dim useSeed As Long
    
    EnsureSettingsSheet
    
    settingsNum = 0
    On Error Resume Next
    If SheetExists(SHEET_SETTINGS) Then
        settingsNum = CLng(ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B2").Value)
    End If
    On Error GoTo 0
    
    trackerMax = 0
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then GoTo BuildNumber
    
    Set wb = OpenTrackingWorkbook(trackingPath, True)
    If wb Is Nothing Then GoTo BuildNumber
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wb) Then
        CloseTrackingWorkbook wb, False
        GoTo BuildNumber
    End If
    
    Set ws = wb.Sheets(TRACK_SHEET_INVOICES)
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        invNum = ExtractSequence(Trim(CStr(ws.Cells(i, COL_I_NUMBER).Value)), INVOICE_PREFIX)
        If invNum > trackerMax Then trackerMax = invNum
    Next i
    
    CloseTrackingWorkbook wb, False

BuildNumber:
    If trackerMax > settingsNum Then
        If Not mWarnedInvoiceCounter Then
            mWarnedInvoiceCounter = True
            Dim msg As String
            msg = "Settings!B2 says next " & INVOICE_PREFIX & " invoice is " & (settingsNum + 1) & "," & vbNewLine & _
                  "but the tracker already contains " & INVOICE_PREFIX & "-" & Format(trackerMax, "0000") & "." & vbNewLine & vbNewLine & _
                  "To avoid duplicate invoice numbers, the next invoice will be " & INVOICE_PREFIX & "-" & Format(trackerMax + 1, "0000") & "." & vbNewLine & vbNewLine & _
                  "(This warning will not repeat in this session.)" & vbNewLine & vbNewLine & _
                  "If you want to RESET numbering, set Settings!B2 to your desired starting number and " & _
                  "remove higher-numbered " & INVOICE_PREFIX & "-### rows from the tracker."
            ShowWarning msg, "Invoice Number Auto-Correction"
        End If
        useSeed = trackerMax
    Else
        useSeed = settingsNum
    End If
    
    GetNextInvoiceNumberFromTracking = INVOICE_PREFIX & "-" & Format(useSeed + 1, "0000")
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then CloseTrackingWorkbook wb, False
    
    If settingsNum > 0 Then
        GetNextInvoiceNumberFromTracking = INVOICE_PREFIX & "-" & Format(settingsNum + 1, "0000")
    Else
        GetNextInvoiceNumberFromTracking = "ERROR"
    End If
End Function

Public Sub PersistQuoteCounter(ByVal quoteNum As String)
    Dim seq As Long
    
    EnsureSettingsSheet
    
    On Error Resume Next
    seq = ExtractSequence(quoteNum, QUOTE_PREFIX)
    If seq > 0 And SheetExists(SHEET_SETTINGS) Then
        ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B1").Value = seq
    End If
    On Error GoTo 0
End Sub

Public Sub PersistInvoiceCounter(ByVal invoiceNum As String)
    Dim seq As Long
    
    EnsureSettingsSheet
    
    On Error Resume Next
    seq = ExtractSequence(invoiceNum, INVOICE_PREFIX)
    If seq > 0 And SheetExists(SHEET_SETTINGS) Then
        ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B2").Value = seq
    End If
    On Error GoTo 0
End Sub

' =====================================================
' ===== APPLY CREDIT TO INVOICE =====
' =====================================================

Public Sub ApplyCreditToInvoice(ByVal invoiceNum As String, ByVal creditAmount As Double, ByVal reason As String)
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    Dim existingCredit As Double, newCredit As Double
    Dim existingReason As String
    
    If Trim(invoiceNum) = "" Then
        ShowError "Invoice number is empty."
        Exit Sub
    End If
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath)
    If wbTrack Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        ShowError "Invoice tracking sheet not found."
        Exit Sub
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    targetRow = FindDocumentRow(wsTrack, invoiceNum, COL_I_NUMBER)
    
    If targetRow = 0 Then
        CloseTrackingWorkbook wbTrack, False
        ShowError "Invoice " & invoiceNum & " not found in tracking."
        Exit Sub
    End If
    
    existingCredit = r2(wsTrack.Cells(targetRow, COL_I_CREDITAMT).Value)
    newCredit = Round(existingCredit + Round(creditAmount, 2), 2)
    
    wsTrack.Cells(targetRow, COL_I_CREDITAMT).Value = newCredit
    wsTrack.Cells(targetRow, COL_I_CREDITAMT).NumberFormat = "#,##0.00"
    
    existingReason = Trim(CStr(wsTrack.Cells(targetRow, COL_I_CREDITREASON).Value))
    If existingReason <> "" Then
        wsTrack.Cells(targetRow, COL_I_CREDITREASON).Value = _
            existingReason & " | " & reason & " (" & FmtR(creditAmount) & ")"
    Else
        wsTrack.Cells(targetRow, COL_I_CREDITREASON).Value = _
            reason & " (" & FmtR(creditAmount) & ")"
    End If
    
    wsTrack.Cells(targetRow, COL_I_LASTMOD).Value = Now
    
    CloseTrackingWorkbook wbTrack, True
    
    ShowInfo "Credit of " & FmtR(creditAmount) & " applied to " & invoiceNum & "." & vbNewLine & _
             "Reason: " & reason & vbNewLine & _
             "Total credit on this invoice: " & FmtR(newCredit), _
             "Credit Applied"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ShowError "Error applying credit: " & Err.Description
End Sub

Public Sub ShowApplyCreditForm()
    Dim invoiceNum As String, reason As String
    Dim amtText As String, amt As Double
    
    invoiceNum = Trim(InputBox("Enter the Invoice Number to credit:", "Apply Credit"))
    If invoiceNum = "" Then Exit Sub
    
    amtText = Trim(InputBox("Enter credit amount (e.g. 150.00):", "Apply Credit - " & invoiceNum))
    If amtText = "" Then Exit Sub
    
    If Not IsNumeric(amtText) Then
        ShowError "Invalid amount: " & amtText
        Exit Sub
    End If
    
    amt = r2(amtText)
    If amt <= 0 Then
        ShowError "Credit amount must be greater than zero."
        Exit Sub
    End If
    
    reason = Trim(InputBox("Reason for credit:", "Apply Credit - " & invoiceNum))
    If reason = "" Then reason = "Credit issued"
    
    ApplyCreditToInvoice invoiceNum, amt, reason
End Sub

' =====================================================
' ===== TRACK QUOTE =====
' =====================================================

Public Sub TrackQuote( _
    ByVal quoteNum As String, _
    ByVal quoteDate As Date, _
    ByVal customerName As String, _
    ByVal patientName As String, _
    ByVal applianceType As String, _
    ByVal subTotal As Variant, _
    ByVal discount As Variant, _
    ByVal vat As Variant, _
    ByVal totalIncl As Variant, _
    ByVal status As String, _
    Optional ByVal invoiceNum As String = "", _
    Optional ByVal filePath As String = "")
    
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    
    If Trim(quoteNum) = "" Then
        ShowError "Cannot track quote: Quote number is empty."
        Exit Sub
    End If
    
    trackingPath = GetQuoteTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath)
    If wbTrack Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_QUOTES, wbTrack) Then
        Set wsTrack = wbTrack.Sheets.Add(After:=wbTrack.Sheets(wbTrack.Sheets.count))
        wsTrack.Name = TRACK_SHEET_QUOTES
        CreateQuoteTrackingHeaders wsTrack
    Else
        Set wsTrack = wbTrack.Sheets(TRACK_SHEET_QUOTES)
    End If
    
    targetRow = FindDocumentRow(wsTrack, quoteNum, COL_Q_NUMBER)
    
    If targetRow = 0 Then
        targetRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row + 1
        If targetRow < 2 Then targetRow = 2
    End If
    
    With wsTrack
        .Cells(targetRow, COL_Q_NUMBER).Value = quoteNum
        .Cells(targetRow, COL_Q_DATE).Value = quoteDate
        .Cells(targetRow, COL_Q_CUSTOMER).Value = customerName
        .Cells(targetRow, COL_Q_PATIENT).Value = patientName
        .Cells(targetRow, COL_Q_APPLIANCE).Value = applianceType
        .Cells(targetRow, COL_Q_SUBTOTAL).Value = r2(subTotal)
        .Cells(targetRow, COL_Q_DISCOUNT).Value = r2(discount)
        .Cells(targetRow, COL_Q_VAT).Value = r2(vat)
        .Cells(targetRow, COL_Q_TOTAL).Value = r2(totalIncl)
        .Cells(targetRow, COL_Q_STATUS).Value = status
        .Cells(targetRow, COL_Q_INVOICENUM).Value = invoiceNum
        .Cells(targetRow, COL_Q_FILEPATH).Value = filePath
        .Cells(targetRow, COL_Q_LASTMOD).Value = Now
        
        .Cells(targetRow, COL_Q_SUBTOTAL).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_Q_DISCOUNT).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_Q_VAT).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_Q_TOTAL).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_Q_DATE).NumberFormat = "YYYY/MM/DD"
        .Cells(targetRow, COL_Q_LASTMOD).NumberFormat = "YYYY/MM/DD HH:MM"
    End With
    
    CloseTrackingWorkbook wbTrack, True
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ShowError "Error tracking quote: " & Err.Description
End Sub

Public Sub UpdateQuoteStatus(ByVal quoteNum As String, ByVal newStatus As String, Optional ByVal invoiceNum As String = "")
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    
    If Trim(quoteNum) = "" Then Exit Sub
    
    trackingPath = GetQuoteTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath)
    If wbTrack Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_QUOTES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        Exit Sub
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_QUOTES)
    targetRow = FindDocumentRow(wsTrack, quoteNum, COL_Q_NUMBER)
    
    If targetRow > 0 Then
        wsTrack.Cells(targetRow, COL_Q_STATUS).Value = newStatus
        If invoiceNum <> "" Then
            wsTrack.Cells(targetRow, COL_Q_INVOICENUM).Value = invoiceNum
        End If
        wsTrack.Cells(targetRow, COL_Q_LASTMOD).Value = Now
    End If
    
    CloseTrackingWorkbook wbTrack, True
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ShowError "Error updating quote status: " & Err.Description
End Sub

' =====================================================
' ===== TRACK INVOICE =====
' =====================================================

Public Sub TrackInvoice( _
    ByVal invoiceNum As String, _
    Optional ByVal quoteNum As String = "", _
    Optional ByVal invoiceDate As Variant, _
    Optional ByVal customerName As String = "", _
    Optional ByVal patientName As String = "", _
    Optional ByVal applianceType As String = "", _
    Optional ByVal subTotal As Variant, _
    Optional ByVal discount As Variant, _
    Optional ByVal vat As Variant, _
    Optional ByVal totalIncl As Variant, _
    Optional ByVal status As String = "Unpaid", _
    Optional ByVal paidDate As Variant, _
    Optional ByVal filePath As String = "")
    
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    
    If Trim(invoiceNum) = "" Then
        ShowError "Cannot track invoice: Invoice number is empty."
        Exit Sub
    End If
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath)
    If wbTrack Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        Set wsTrack = wbTrack.Sheets.Add(After:=wbTrack.Sheets(wbTrack.Sheets.count))
        wsTrack.Name = TRACK_SHEET_INVOICES
        CreateInvoiceTrackingHeaders wsTrack
    Else
        Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
        EnsurePaidAmountHeader wsTrack
    End If
    
    targetRow = FindDocumentRow(wsTrack, invoiceNum, COL_I_NUMBER)
    
    If targetRow = 0 Then
        targetRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row + 1
        If targetRow < 2 Then targetRow = 2
    End If
    
    With wsTrack
        .Cells(targetRow, COL_I_NUMBER).Value = invoiceNum
        .Cells(targetRow, COL_I_QUOTENUM).Value = quoteNum
        
        If Not IsMissing(invoiceDate) Then
            If IsDate(invoiceDate) Then
                .Cells(targetRow, COL_I_DATE).Value = CDate(invoiceDate)
            End If
        End If
        
        .Cells(targetRow, COL_I_CUSTOMER).Value = customerName
        .Cells(targetRow, COL_I_PATIENT).Value = patientName
        .Cells(targetRow, COL_I_APPLIANCE).Value = applianceType
        .Cells(targetRow, COL_I_SUBTOTAL).Value = r2(subTotal)
        .Cells(targetRow, COL_I_DISCOUNT).Value = r2(discount)
        .Cells(targetRow, COL_I_VAT).Value = r2(vat)
        .Cells(targetRow, COL_I_TOTAL).Value = r2(totalIncl)
        .Cells(targetRow, COL_I_STATUS).Value = status
        
        If Not IsMissing(paidDate) Then
            If IsDate(paidDate) Then
                .Cells(targetRow, COL_I_PAIDDATE).Value = paidDate
            End If
        End If
        
        .Cells(targetRow, COL_I_FILEPATH).Value = filePath
        .Cells(targetRow, COL_I_LASTMOD).Value = Now
        
        .Cells(targetRow, COL_I_SUBTOTAL).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_I_DISCOUNT).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_I_VAT).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_I_TOTAL).NumberFormat = "R#,##0.00"
        .Cells(targetRow, COL_I_DATE).NumberFormat = "YYYY/MM/DD"
        .Cells(targetRow, COL_I_PAIDDATE).NumberFormat = "YYYY/MM/DD"
        .Cells(targetRow, COL_I_LASTMOD).NumberFormat = "YYYY/MM/DD HH:MM"
    End With
    
    CloseTrackingWorkbook wbTrack, True
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ShowError "Error tracking invoice: " & Err.Description
End Sub

' =====================================================
' ===== UPDATE INVOICE STATUS =====
' =====================================================

Public Sub UpdateInvoiceStatus(ByVal invoiceNum As String, ByVal newStatus As String, Optional ByVal paidDate As Variant)
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim targetRow As Long, trackingPath As String
    Dim total As Double, credit As Double, due As Double
    
    If Trim(invoiceNum) = "" Then Exit Sub
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath)
    If wbTrack Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        Exit Sub
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    
    EnsurePaidAmountHeader wsTrack
    
    targetRow = FindDocumentRow(wsTrack, invoiceNum, COL_I_NUMBER)
    
    If targetRow > 0 Then
        wsTrack.Cells(targetRow, COL_I_STATUS).Value = newStatus
        
        If Not IsMissing(paidDate) Then
            If IsDate(paidDate) Then
                wsTrack.Cells(targetRow, COL_I_PAIDDATE).Value = paidDate
            End If
        End If
        
        If LCase$(newStatus) = LCase$(STATUS_INVOICE_PAID) Then
            total = r2(wsTrack.Cells(targetRow, COL_I_TOTAL).Value)
            credit = r2(wsTrack.Cells(targetRow, COL_I_CREDITAMT).Value)
            due = Round(total - credit, 2)
            If due < 0 Then due = 0
            wsTrack.Cells(targetRow, COL_I_PAIDAMOUNT).Value = due
            wsTrack.Cells(targetRow, COL_I_PAIDAMOUNT).NumberFormat = "R#,##0.00"
        ElseIf LCase$(newStatus) = LCase$(STATUS_INVOICE_UNPAID) Then
            wsTrack.Cells(targetRow, COL_I_PAIDAMOUNT).Value = 0
            wsTrack.Cells(targetRow, COL_I_PAIDAMOUNT).NumberFormat = "R#,##0.00"
        End If
        
        wsTrack.Cells(targetRow, COL_I_LASTMOD).Value = Now
    End If
    
    CloseTrackingWorkbook wbTrack, True
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    ShowError "Error updating invoice status: " & Err.Description
End Sub

' =====================================================
' ===== SYNC QUOTE FROM FILE =====
' =====================================================

Public Sub SyncQuoteFromFile()
    Dim filePath As String, wbExport As Workbook, wsQuote As Worksheet
    Dim quoteNum As String
    Dim vSubTotal As Double, vDiscount As Double, vVAT As Double, vTotal As Double
    
    filePath = GetOpenPath("Excel Files (*.xls*), *.xls*", "Select Exported Quote File to Sync")
    If filePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Syncing quote..."
    
    Set wbExport = Workbooks.Open(filePath, UpdateLinks:=False, readOnly:=True)
    
    If Not SheetExistsInWorkbook(SHEET_QUOTE, wbExport) Then
        ShowError "Quote sheet not found in the selected file."
        wbExport.Close saveChanges:=False
        GoTo CleanUp
    End If
    
    Set wsQuote = wbExport.Sheets(SHEET_QUOTE)
    quoteNum = SafeString(wsQuote.Range(CELL_DOC_NUMBER).Value)
    
    If Trim(quoteNum) = "" Then
        ShowError "No quote number found in the file."
        wbExport.Close saveChanges:=False
        GoTo CleanUp
    End If
    
    vSubTotal = r2(wsQuote.Range(CELL_SUMMARY_SUBTOT).Value)
    vDiscount = r2(wsQuote.Range(CELL_SUMMARY_DISC).Value)
    vVAT = r2(wsQuote.Range(CELL_SUMMARY_VAT).Value)
    vTotal = r2(wsQuote.Range(CELL_SUMMARY_TOTAL).Value)
    
    TrackQuote quoteNum, _
               SafeDate(wsQuote.Range(CELL_DOC_DATE).Value), _
               SafeString(wsQuote.Range(CELL_CUSTOMER).Value), _
               SafeString(wsQuote.Range(CELL_PATIENT).Value), _
               SafeString(wsQuote.Range(CELL_APPLIANCE).Value), _
               vSubTotal, vDiscount, vVAT, vTotal, _
               STATUS_QUOTE_OPEN, "", filePath
    
    wbExport.Close saveChanges:=False
    
    RestoreExcelState
    ShowInfo "Quote " & quoteNum & " synced successfully!", "Sync Complete"
    Exit Sub

CleanUp:
    RestoreExcelState
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close saveChanges:=False
    RestoreExcelState
    ShowError "Error syncing quote: " & Err.Description
End Sub

' =====================================================
' ===== SYNC INVOICE FROM FILE =====
' =====================================================

Public Sub SyncInvoiceFromFile()
    Dim filePath As String, wbExport As Workbook, wsInvoice As Worksheet
    Dim invoiceNum As String, quoteNum As String
    Dim vSubTotal As Double, vDiscount As Double, vVAT As Double, vTotal As Double
    
    filePath = GetOpenPath("Excel Files (*.xls*), *.xls*", "Select Exported Invoice File to Sync")
    If filePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Syncing invoice..."
    
    Set wbExport = Workbooks.Open(filePath, UpdateLinks:=False, readOnly:=True)
    
    If Not SheetExistsInWorkbook(SHEET_INVOICE, wbExport) Then
        ShowError "Invoice sheet not found in the selected file."
        wbExport.Close saveChanges:=False
        GoTo CleanUp
    End If
    
    Set wsInvoice = wbExport.Sheets(SHEET_INVOICE)
    invoiceNum = SafeString(wsInvoice.Range(CELL_DOC_NUMBER).Value)
    
    If Trim(invoiceNum) = "" Then
        ShowError "No invoice number found in the file."
        wbExport.Close saveChanges:=False
        GoTo CleanUp
    End If
    
    ' Quote number is no longer on the invoice sheet.
    ' Preserve any existing link in the tracker.
    quoteNum = LookupExistingInvoiceQuoteNum(invoiceNum)
    
    vSubTotal = r2(wsInvoice.Range(CELL_SUMMARY_SUBTOT).Value)
    vDiscount = r2(wsInvoice.Range(CELL_SUMMARY_DISC).Value)
    vVAT = r2(wsInvoice.Range(CELL_SUMMARY_VAT).Value)
    vTotal = r2(wsInvoice.Range(CELL_SUMMARY_TOTAL).Value)
    
    TrackInvoice invoiceNum, quoteNum, _
                 SafeDate(wsInvoice.Range(CELL_DOC_DATE).Value), _
                 SafeString(wsInvoice.Range(CELL_CUSTOMER).Value), _
                 SafeString(wsInvoice.Range(CELL_PATIENT).Value), _
                 SafeString(wsInvoice.Range(CELL_APPLIANCE).Value), _
                 vSubTotal, vDiscount, vVAT, vTotal, _
                 STATUS_INVOICE_UNPAID, Empty, filePath
    
    wbExport.Close saveChanges:=False
    
    RestoreExcelState
    ShowInfo "Invoice " & invoiceNum & " synced successfully!", "Sync Complete"
    Exit Sub

CleanUp:
    RestoreExcelState
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close saveChanges:=False
    RestoreExcelState
    ShowError "Error syncing invoice: " & Err.Description
End Sub

' =====================================================
' ===== READ INVOICE TOTALS (for forms) =====
' =====================================================

Public Sub ReadInvoiceTotals(ByVal invoiceNum As String, _
                             ByRef total As Double, _
                             ByRef paid As Double, _
                             ByRef credit As Double)
    Dim rec As InvoiceRecord
    rec = ReadInvoiceRecord(invoiceNum)
    total = rec.totalIncl
    paid = rec.paidAmount
    credit = rec.creditAmount
End Sub

' =====================================================
' ===== LOOKUP EXISTING INVOICE -> QUOTE LINK =====
' =====================================================
' Returns the QuoteNumber already stored against this invoice in the tracker,
' or "" if there isn't one. Used so re-syncing/updating an invoice doesn't
' blank the quote-number link (since quote # is no longer on the invoice sheet).

Public Function LookupExistingInvoiceQuoteNum(ByVal invoiceNum As String) As String
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim trackingPath As String, targetRow As Long
    
    LookupExistingInvoiceQuoteNum = ""
    
    If Trim(invoiceNum) = "" Then Exit Function
    
    On Error GoTo SafeExit
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then GoTo SafeExit
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then GoTo SafeExit
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        GoTo SafeExit
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    targetRow = FindDocumentRow(wsTrack, invoiceNum, COL_I_NUMBER)
    
    If targetRow > 0 Then
        LookupExistingInvoiceQuoteNum = SafeString(wsTrack.Cells(targetRow, COL_I_QUOTENUM).Value)
    End If
    
    CloseTrackingWorkbook wbTrack, False
    Exit Function

SafeExit:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
End Function

