Attribute VB_Name = "modTrackerSafety"
Option Explicit

' =====================================================
' ===== TRACKER WRITE SAFETY — verify, retry, audit =====
' Wraps TrackInvoice / TrackQuote with a re-read verification
' pass so the user never sees "success" for a write that
' silently failed to reach disk.
' =====================================================

' Master switch. Set to False to disable verification entirely.
Public Const VERIFY_TRACKER_WRITES As Boolean = True
' How many times to retry a failed write before giving up.
Public Const TRACKER_RETRY_COUNT   As Long = 2
' Hidden audit log sheet name.
Public Const TRACKER_AUDIT_SHEET   As String = "AuditLog"
' Small pause between retries so transient locks (OneDrive sync, etc.) can clear.
Private Const TRACKER_RETRY_PAUSE_MS As Long = 250

' =====================================================
' ===== PUBLIC API =====
' =====================================================

' Returns True if the invoice was successfully written AND verified.
Public Function SaveAndVerifyInvoice( _
    ByVal invoiceNum As String, _
    ByVal quoteNum As String, _
    ByVal invoiceDate As Date, _
    ByVal customerName As String, _
    ByVal patientName As String, _
    ByVal applianceType As String, _
    ByVal subTotal As Variant, _
    ByVal discount As Variant, _
    ByVal vat As Variant, _
    ByVal totalIncl As Variant, _
    ByVal status As String, _
    ByVal paidDate As Variant, _
    ByVal filePath As String) As Boolean
    
    Dim attempt As Long
    Dim writeOk As Boolean, verifyOk As Boolean
    Dim mismatch As String
    Dim expectedTotal As Double
    Dim resultText As String
    
    expectedTotal = r2(totalIncl)
    
    For attempt = 0 To TRACKER_RETRY_COUNT
        ' --- write ---
        writeOk = TryTrackInvoice(invoiceNum, quoteNum, invoiceDate, _
                                   customerName, patientName, applianceType, _
                                   subTotal, discount, vat, totalIncl, _
                                   status, paidDate, filePath)
        
        If Not writeOk Then
            LogAudit "WriteInvoice", invoiceNum, customerName, patientName, expectedTotal, _
                     "FAILED-WRITE (attempt " & (attempt + 1) & ")", _
                     "TrackInvoice raised an error"
            GoTo NextAttempt
        End If
        
        ' --- verify ---
        If Not VERIFY_TRACKER_WRITES Then
            LogAudit "WriteInvoice", invoiceNum, customerName, patientName, expectedTotal, _
                     "OK (no verify)", "Verification disabled"
            SaveAndVerifyInvoice = True
            Exit Function
        End If
        
        verifyOk = VerifyInvoiceRow(invoiceNum, expectedTotal, mismatch)
        
        If verifyOk Then
            If attempt = 0 Then
                resultText = "OK"
            Else
                resultText = "OK (retry " & attempt & ")"
            End If
            LogAudit "WriteInvoice", invoiceNum, customerName, patientName, expectedTotal, _
                     resultText, "Verified in tracker"
            SaveAndVerifyInvoice = True
            Exit Function
        End If
        
        LogAudit "WriteInvoice", invoiceNum, customerName, patientName, expectedTotal, _
                 "FAILED-VERIFY (attempt " & (attempt + 1) & ")", mismatch
NextAttempt:
        DelayMs TRACKER_RETRY_PAUSE_MS
    Next attempt
    
    SaveAndVerifyInvoice = False
End Function

' Returns True if the quote was successfully written AND verified.
Public Function SaveAndVerifyQuote( _
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
    ByVal invoiceNum As String, _
    ByVal filePath As String) As Boolean
    
    Dim attempt As Long
    Dim writeOk As Boolean, verifyOk As Boolean
    Dim mismatch As String
    Dim expectedTotal As Double
    Dim resultText As String
    
    expectedTotal = r2(totalIncl)
    
    For attempt = 0 To TRACKER_RETRY_COUNT
        writeOk = TryTrackQuote(quoteNum, quoteDate, customerName, patientName, applianceType, _
                                 subTotal, discount, vat, totalIncl, status, invoiceNum, filePath)
        
        If Not writeOk Then
            LogAudit "WriteQuote", quoteNum, customerName, patientName, expectedTotal, _
                     "FAILED-WRITE (attempt " & (attempt + 1) & ")", _
                     "TrackQuote raised an error"
            GoTo NextAttempt
        End If
        
        If Not VERIFY_TRACKER_WRITES Then
            LogAudit "WriteQuote", quoteNum, customerName, patientName, expectedTotal, _
                     "OK (no verify)", "Verification disabled"
            SaveAndVerifyQuote = True
            Exit Function
        End If
        
        verifyOk = VerifyQuoteRow(quoteNum, expectedTotal, mismatch)
        
        If verifyOk Then
            If attempt = 0 Then
                resultText = "OK"
            Else
                resultText = "OK (retry " & attempt & ")"
            End If
            LogAudit "WriteQuote", quoteNum, customerName, patientName, expectedTotal, _
                     resultText, "Verified in tracker"
            SaveAndVerifyQuote = True
            Exit Function
        End If
        
        LogAudit "WriteQuote", quoteNum, customerName, patientName, expectedTotal, _
                 "FAILED-VERIFY (attempt " & (attempt + 1) & ")", mismatch
NextAttempt:
        DelayMs TRACKER_RETRY_PAUSE_MS
    Next attempt
    
    SaveAndVerifyQuote = False
End Function

' Shows a green confirmation popup. Call after SaveAndVerify* returned True.
Public Sub ShowVerifiedConfirmation(ByVal docKind As String, ByVal docNum As String, _
                                    ByVal customer As String, ByVal patient As String, _
                                    ByVal totalIncl As Double, ByVal extraLine As String)
    Dim msg As String
    msg = ChrW(10004) & " " & docKind & " saved successfully" & vbNewLine & vbNewLine & _
          docKind & " : " & docNum & vbNewLine & _
          "Customer: " & customer & vbNewLine & _
          "Patient : " & patient & vbNewLine & _
          "Total   : " & FmtR(totalIncl) & vbNewLine & _
          "Status  : Verified in tracker " & ChrW(10004)
    
    If Len(extraLine) > 0 Then msg = msg & vbNewLine & vbNewLine & extraLine
    
    MsgBox msg, vbInformation, docKind & " Saved"
End Sub

' Shows a red error popup. Call after SaveAndVerify* returned False.
Public Sub ShowVerifyFailure(ByVal docKind As String, ByVal docNum As String, _
                             ByVal customer As String, ByVal patient As String, _
                             ByVal totalIncl As Double)
    Dim msg As String
    msg = ChrW(10006) & " " & docKind & " was NOT saved to the tracker." & vbNewLine & vbNewLine & _
          docKind & " : " & docNum & vbNewLine & _
          "Customer: " & customer & vbNewLine & _
          "Patient : " & patient & vbNewLine & _
          "Total   : " & FmtR(totalIncl) & vbNewLine & vbNewLine & _
          "The " & LCase$(docKind) & " sheet was filled in, but the row" & vbNewLine & _
          "could NOT be confirmed in the tracking file." & vbNewLine & vbNewLine & _
          "Statements and the dashboard will MISS it until fixed." & vbNewLine & vbNewLine & _
          "Possible causes:" & vbNewLine & _
          "  - Tracking file is open in another window or by another user" & vbNewLine & _
          "  - OneDrive / network share is offline or syncing" & vbNewLine & _
          "  - File is read-only or locked" & vbNewLine & vbNewLine & _
          "See the AuditLog sheet (hidden) for details, then click" & vbNewLine & _
          "'" & docKind & " Update to Tracker' to retry."
    
    MsgBox msg, vbCritical, "Save Failed - " & docKind
End Sub

' =====================================================
' ===== INTERNAL - VERIFICATION =====
' =====================================================

Private Function VerifyInvoiceRow(ByVal invoiceNum As String, _
                                   ByVal expectedTotal As Double, _
                                   ByRef mismatchOut As String) As Boolean
    Dim rec As InvoiceRecord
    Dim diff As Double
    
    mismatchOut = ""
    
    On Error GoTo Failed
    rec = ReadInvoiceRecord(invoiceNum)
    
    If Not rec.found Then
        mismatchOut = "Row not found in tracker after write"
        VerifyInvoiceRow = False
        Exit Function
    End If
    
    diff = Round(rec.totalIncl - expectedTotal, 2)
    If Abs(diff) > 0.01 Then
        mismatchOut = "Row found but Total = " & FmtR(rec.totalIncl) & _
                      " (expected " & FmtR(expectedTotal) & ")"
        VerifyInvoiceRow = False
        Exit Function
    End If
    
    VerifyInvoiceRow = True
    Exit Function

Failed:
    mismatchOut = "Verification raised an error: " & Err.Description
    VerifyInvoiceRow = False
End Function

Private Function VerifyQuoteRow(ByVal quoteNum As String, _
                                 ByVal expectedTotal As Double, _
                                 ByRef mismatchOut As String) As Boolean
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim trackingPath As String, targetRow As Long
    Dim foundTotal As Double, diff As Double
    
    mismatchOut = ""
    
    On Error GoTo Failed
    
    trackingPath = GetQuoteTrackingPath()
    If trackingPath = "" Then
        mismatchOut = "Quote tracking path not set"
        VerifyQuoteRow = False
        Exit Function
    End If
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then
        mismatchOut = "Cannot reopen quote tracker for verification"
        VerifyQuoteRow = False
        Exit Function
    End If
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_QUOTES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        mismatchOut = "Quotes sheet missing from tracker"
        VerifyQuoteRow = False
        Exit Function
    End If
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_QUOTES)
    targetRow = FindDocumentRow(wsTrack, quoteNum, COL_Q_NUMBER)
    
    If targetRow = 0 Then
        CloseTrackingWorkbook wbTrack, False
        mismatchOut = "Row not found in tracker after write"
        VerifyQuoteRow = False
        Exit Function
    End If
    
    foundTotal = r2(wsTrack.Cells(targetRow, COL_Q_TOTAL).Value)
    CloseTrackingWorkbook wbTrack, False
    
    diff = Round(foundTotal - expectedTotal, 2)
    If Abs(diff) > 0.01 Then
        mismatchOut = "Row found but Total = " & FmtR(foundTotal) & _
                      " (expected " & FmtR(expectedTotal) & ")"
        VerifyQuoteRow = False
        Exit Function
    End If
    
    VerifyQuoteRow = True
    Exit Function

Failed:
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    mismatchOut = "Verification raised an error: " & Err.Description
    VerifyQuoteRow = False
End Function

' =====================================================
' ===== INTERNAL - SAFE CALL WRAPPERS =====
' =====================================================
' These wrappers catch any unhandled error from TrackInvoice / TrackQuote
' and turn it into a Boolean return value instead of a popup.

Private Function TryTrackInvoice(ByVal invoiceNum As String, _
                                  ByVal quoteNum As String, _
                                  ByVal invoiceDate As Date, _
                                  ByVal customerName As String, _
                                  ByVal patientName As String, _
                                  ByVal applianceType As String, _
                                  ByVal subTotal As Variant, _
                                  ByVal discount As Variant, _
                                  ByVal vat As Variant, _
                                  ByVal totalIncl As Variant, _
                                  ByVal status As String, _
                                  ByVal paidDate As Variant, _
                                  ByVal filePath As String) As Boolean
    On Error GoTo Failed
    TrackInvoice invoiceNum, quoteNum, invoiceDate, customerName, patientName, applianceType, _
                 subTotal, discount, vat, totalIncl, status, paidDate, filePath
    TryTrackInvoice = True
    Exit Function
Failed:
    TryTrackInvoice = False
End Function

Private Function TryTrackQuote(ByVal quoteNum As String, _
                                ByVal quoteDate As Date, _
                                ByVal customerName As String, _
                                ByVal patientName As String, _
                                ByVal applianceType As String, _
                                ByVal subTotal As Variant, _
                                ByVal discount As Variant, _
                                ByVal vat As Variant, _
                                ByVal totalIncl As Variant, _
                                ByVal status As String, _
                                ByVal invoiceNum As String, _
                                ByVal filePath As String) As Boolean
    On Error GoTo Failed
    TrackQuote quoteNum, quoteDate, customerName, patientName, applianceType, _
               subTotal, discount, vat, totalIncl, status, invoiceNum, filePath
    TryTrackQuote = True
    Exit Function
Failed:
    TryTrackQuote = False
End Function

' =====================================================
' ===== AUDIT LOG =====
' =====================================================
' Important: errors during audit-log writes must NOT bubble up to callers
' (that would mask the real error they're trying to log). But we also
' must not blanket-swallow errors so widely that we hide bugs forever.
'
' Strategy:
'   - Look up / create the sheet WITHIN a narrow On Error Resume Next.
'   - As soon as we have a valid sheet, switch to On Error GoTo SafeExit
'     so any write failures land in a single trap that just exits silently.

Public Sub LogAudit(ByVal action As String, ByVal docNum As String, _
                    ByVal customer As String, ByVal patient As String, _
                    ByVal total As Double, ByVal result As String, _
                    ByVal detail As String)
    Dim ws As Worksheet, nextRow As Long
    
    ' --- Step 1: locate (or create) the audit-log sheet ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TRACKER_AUDIT_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = EnsureAuditLogSheet()
        If ws Is Nothing Then Exit Sub
    End If
    
    ' --- Step 2: tight error trap from here on ---
    On Error GoTo SafeExit
    
    nextRow = ws.Cells(ws.rows.count, "A").End(xlUp).row + 1
    If nextRow < 2 Then nextRow = 2
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "YYYY/MM/DD HH:MM:SS"
    ws.Cells(nextRow, 2).Value = action
    ws.Cells(nextRow, 3).Value = docNum
    ws.Cells(nextRow, 4).Value = customer
    ws.Cells(nextRow, 5).Value = patient
    ws.Cells(nextRow, 6).Value = total
    ws.Cells(nextRow, 6).NumberFormat = "R#,##0.00"
    ws.Cells(nextRow, 7).Value = result
    ws.Cells(nextRow, 8).Value = detail
    
    ' Colour code result column for at-a-glance scanning
    Select Case True
        Case InStr(1, result, "OK", vbTextCompare) > 0
            ws.Cells(nextRow, 7).Interior.Color = RGB(220, 245, 220) ' light green
        Case InStr(1, result, "FAILED", vbTextCompare) > 0
            ws.Cells(nextRow, 7).Interior.Color = RGB(255, 220, 220) ' light red
        Case InStr(1, result, "OVERPAYMENT", vbTextCompare) > 0
            ws.Cells(nextRow, 7).Interior.Color = RGB(255, 235, 195) ' deeper amber for overpayments
        Case Else
            ws.Cells(nextRow, 7).Interior.Color = RGB(255, 245, 200) ' light amber
    End Select
    
SafeExit:
End Sub

Private Function EnsureAuditLogSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TRACKER_AUDIT_SHEET)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Set EnsureAuditLogSheet = ws
        Exit Function
    End If
    
    On Error GoTo Failed
    
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    Application.DisplayAlerts = True
    
    ws.Name = TRACKER_AUDIT_SHEET
    ws.Visible = xlSheetVeryHidden   ' very hidden - can't be unhidden from the GUI
    
    ws.Range("A1:H1").Value = Array("Timestamp", "Action", "DocNumber", "Customer", _
                                     "Patient", "Total", "Result", "Detail")
    With ws.Range("A1:H1")
        .Font.bold = True
        .Interior.Color = RGB(46, 135, 176)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns("A:H").AutoFit
    
    Set EnsureAuditLogSheet = ws
    Exit Function

Failed:
    Application.DisplayAlerts = True
    Set EnsureAuditLogSheet = Nothing
End Function

' Lightweight sleep without the Sleep API (cross-platform).
Private Sub DelayMs(ByVal ms As Long)
    Dim t As Double
    t = Timer + ms / 1000#
    Do While Timer < t
        DoEvents
    Loop
End Sub

' =====================================================
' ===== USER-INVOKABLE - REVEAL / HIDE AUDIT LOG =====
' Run from Alt+F8 if you need to inspect the log.
' =====================================================

Public Sub ShowAuditLog()
    Dim ws As Worksheet
    Set ws = EnsureAuditLogSheet()
    If ws Is Nothing Then
        MsgBox "Unable to open audit log.", vbCritical, "Audit Log"
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range("A1").Select
End Sub

Public Sub HideAuditLog()
    On Error Resume Next
    ThisWorkbook.Sheets(TRACKER_AUDIT_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
    
    On Error Resume Next
    If SheetExists(SHEET_MENU) Then ThisWorkbook.Sheets(SHEET_MENU).Activate
    On Error GoTo 0
End Sub

' =====================================================
' ===== MAINTENANCE - SCAN FOR ORPHAN INVOICES =====
' For invoices that slipped through BEFORE this safety net.
' =====================================================

Public Sub ScanForOrphanInvoices()
    Dim trackingPath As String, exportFolder As String
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim trackedNums As Object   ' Scripting.Dictionary
    Dim lastRow As Long, i As Long
    Dim f As String, baseName As String, possibleInv As String
    Dim orphans As String, orphanCount As Long
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    exportFolder = InputBox( _
        "Enter the folder where you save exported invoices (e.g. D:\WeDental\Invoices\):", _
        "Scan for Orphan Invoices")
    If exportFolder = "" Then Exit Sub
    
    If Right$(exportFolder, 1) <> "\" Then exportFolder = exportFolder & "\"
    If Not FolderExists(exportFolder) Then
        MsgBox "Folder not found: " & exportFolder, vbExclamation, "Scan"
        Exit Sub
    End If
    
    ' --- Build dictionary of already-tracked invoice numbers ---
    Set trackedNums = CreateTrackedInvoiceDictionary(trackingPath, wbTrack)
    If trackedNums Is Nothing Then Exit Sub   ' caller already saw an error
    
    ' --- Walk the export folder, gathering anything that's missing from the tracker ---
    orphans = ""
    orphanCount = 0
    
    f = Dir(exportFolder & INVOICE_PREFIX & "-*.xlsm")
    Do While f <> ""
        baseName = ExtractInvoiceNumFromFilename(f)
        If baseName <> "" Then
            If Not trackedNums.exists(baseName) Then
                orphans = orphans & baseName & "  (" & f & ")" & vbNewLine
                orphanCount = orphanCount + 1
            End If
        End If
        f = Dir
    Loop
    
    f = Dir(exportFolder & INVOICE_PREFIX & "-*.pdf")
    Do While f <> ""
        baseName = ExtractInvoiceNumFromFilename(f)
        If baseName <> "" Then
            If Not trackedNums.exists(baseName) Then
                orphans = orphans & baseName & "  (" & f & ")" & vbNewLine
                orphanCount = orphanCount + 1
            End If
        End If
        f = Dir
    Loop
    
    If orphanCount = 0 Then
        MsgBox "No orphan invoices found." & vbNewLine & vbNewLine & _
               "Every exported invoice in" & vbNewLine & exportFolder & vbNewLine & _
               "is present in the tracker.", vbInformation, "Scan Complete"
    Else
        MsgBox orphanCount & " orphan invoice(s) found in:" & vbNewLine & _
               exportFolder & vbNewLine & vbNewLine & orphans & vbNewLine & _
               "These invoices were exported but are MISSING from the tracker." & vbNewLine & _
               "Open each one and click 'Update Invoice to Tracker' to recover.", _
               vbExclamation, "Orphan Invoices Found"
    End If
End Sub

' Builds a case-insensitive dictionary of every invoice number currently
' tracked. Caller is responsible for nothing -- the function opens AND
' closes the tracker itself.
Private Function CreateTrackedInvoiceDictionary(ByVal trackingPath As String, _
                                                 ByRef wbTrackOut As Workbook) As Object
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim dict As Object
    Dim lastRow As Long, i As Long
    Dim invNum As String
    
    Set CreateTrackedInvoiceDictionary = Nothing
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then Exit Function
    
    On Error GoTo SafeExit
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' vbTextCompare = case-insensitive
    
    If SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
        lastRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row
        For i = 2 To lastRow
            invNum = Trim(CStr(wsTrack.Cells(i, COL_I_NUMBER).Value))
            If invNum <> "" Then
                If Not dict.exists(invNum) Then dict.Add invNum, True
            End If
        Next i
    End If
    
SafeExit:
    On Error Resume Next
    CloseTrackingWorkbook wbTrack, False
    On Error GoTo 0
    
    Set CreateTrackedInvoiceDictionary = dict
End Function

Private Function ExtractInvoiceNumFromFilename(ByVal f As String) As String
    Dim p As Long, q As Long, s As String, prefix As String
    
    prefix = INVOICE_PREFIX & "-"
    p = InStr(1, f, prefix, vbTextCompare)
    If p = 0 Then
        ExtractInvoiceNumFromFilename = ""
        Exit Function
    End If
    
    s = Mid$(f, p)   ' "INV-0067 something.xlsm"
    q = InStr(1, s, " ")
    If q > 0 Then s = Left$(s, q - 1)
    
    ' Strip extension if no space was present
    q = InStrRev(s, ".")
    If q > 0 Then s = Left$(s, q - 1)
    
    ExtractInvoiceNumFromFilename = s
End Function

