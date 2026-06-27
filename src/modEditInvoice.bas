Attribute VB_Name = "modEditInvoice"
Option Explicit

' =====================================================
' ===== CHANGE INVOICE DATE =====
' Assign this macro to a button on the Invoice sheet:
'   Right-click button -> Assign Macro -> ChangeInvoiceDate
'
' Behaviour:
'   1. Reads the invoice number from the active Invoice sheet (CELL_DOC_NUMBER).
'   2. Opens the invoice tracking workbook.
'   3. Looks up the invoice row by invoice number.
'   4. WARNS if the invoice is already marked as Paid (but allows the change).
'   5. Asks the user for the new date (PickDate, falls back to InputBox).
'   6. Writes the new date back to the tracker and saves.
'   7. Also updates the date on the visible Invoice sheet (CELL_DOC_DATE).
'   8. Audit-logs the change so the date history is recoverable.
' =====================================================
'
' Cell addresses live in modConfig:
'   CELL_DOC_NUMBER = "G8"   invoice number
'   CELL_DOC_DATE   = "G7"   invoice date

Public Sub ChangeInvoiceDate()
    Dim wbTrack As Workbook, wsInv As Worksheet
    Dim trackingPath As String
    Dim invNum As String
    Dim newDate As Date, oldDate As Variant
    Dim foundRow As Long
    Dim status As String, customer As String, patient As String
    Dim total As Double
    Dim resp As VbMsgBoxResult
    Dim defaultDateStr As String
    
    ' --- Step 1: grab the invoice number ---
    invNum = GetActiveInvoiceNumber()
    
    If invNum = "" Then
        invNum = Trim(InputBox( _
            "Enter the Invoice Number to change the date for:" & vbNewLine & vbNewLine & _
            "(e.g. " & INVOICE_PREFIX & "-0067)", _
            "Change Invoice Date"))
    End If
    
    If invNum = "" Then Exit Sub
    
    ' --- Step 2: open tracking workbook ---
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, False)
    If wbTrack Is Nothing Then
        RestoreExcelState
        ShowError "Cannot open invoice tracking file."
        Exit Sub
    End If
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then
        CloseTrackingWorkbook wbTrack, False
        RestoreExcelState
        ShowWarning "No invoice tracking data found.", "No Data"
        Exit Sub
    End If
    
    Set wsInv = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    
    ' --- Step 3: find the invoice row ---
    foundRow = FindInvoiceRowByNumber(wsInv, invNum)
    If foundRow = 0 Then
        CloseTrackingWorkbook wbTrack, False
        RestoreExcelState
        ShowWarning "Invoice '" & invNum & "' was not found in the tracking workbook.", "Not Found"
        Exit Sub
    End If
    
    oldDate = wsInv.Cells(foundRow, COL_I_DATE).Value
    customer = Trim(CStr(wsInv.Cells(foundRow, COL_I_CUSTOMER).Value))
    patient = Trim(CStr(wsInv.Cells(foundRow, COL_I_PATIENT).Value))
    status = Trim(CStr(wsInv.Cells(foundRow, COL_I_STATUS).Value))
    total = r2(wsInv.Cells(foundRow, COL_I_TOTAL).Value)
    
    Application.ScreenUpdating = True
    
    ' --- Step 4: warn if invoice is already paid (but allow) ---
    If LCase$(status) = LCase$(STATUS_INVOICE_PAID) Then
        resp = MsgBox( _
            "Invoice " & invNum & " is marked as PAID." & vbNewLine & vbNewLine & _
            "Customer: " & customer & vbNewLine & _
            "Patient : " & patient & vbNewLine & _
            "Current date: " & FormatDateOrBlank(oldDate) & vbNewLine & vbNewLine & _
            "Are you sure you want to change the date of a paid invoice?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "Paid Invoice")
        If resp <> vbYes Then
            CloseTrackingWorkbook wbTrack, False
            Exit Sub
        End If
    End If
    
    ' --- Step 5: ask for the new date ---
    If IsDate(oldDate) Then
        defaultDateStr = Format(CDate(oldDate), "YYYY/MM/DD")
    Else
        defaultDateStr = Format(Date, "YYYY/MM/DD")
    End If
    
    newDate = TryPickDateOrPrompt(defaultDateStr, invNum, customer, patient)
    If newDate = 0 Then
        CloseTrackingWorkbook wbTrack, False
        Exit Sub
    End If
    
    ' --- Step 6: final confirmation ---
    resp = MsgBox( _
        "Change invoice date as follows?" & vbNewLine & vbNewLine & _
        "Invoice : " & invNum & vbNewLine & _
        "Customer: " & customer & vbNewLine & _
        "Patient : " & patient & vbNewLine & _
        "Old date: " & defaultDateStr & vbNewLine & _
        "New date: " & Format(newDate, "YYYY/MM/DD"), _
        vbQuestion + vbYesNo + vbDefaultButton1, _
        "Confirm Date Change")
    
    If resp <> vbYes Then
        CloseTrackingWorkbook wbTrack, False
        Exit Sub
    End If
    
    ' --- Step 7: write the new date to the tracker ---
    Application.ScreenUpdating = False
    
    wsInv.Cells(foundRow, COL_I_DATE).Value = newDate
    wsInv.Cells(foundRow, COL_I_DATE).NumberFormat = "YYYY/MM/DD"
    wsInv.Cells(foundRow, COL_I_LASTMOD).Value = Now
    wsInv.Cells(foundRow, COL_I_LASTMOD).NumberFormat = "YYYY/MM/DD HH:MM"
    
    CloseTrackingWorkbook wbTrack, True
    
    ' --- Step 8: also update the date on the visible Invoice sheet ---
    TryUpdateInvoiceSheetDate newDate
    
    RestoreExcelState
    
    ' --- Step 9: audit log so the change is traceable later ---
    On Error Resume Next
    LogAudit "ChangeInvoiceDate", invNum, customer, patient, total, _
             "OK", _
             "Date changed from " & defaultDateStr & " to " & Format(newDate, "YYYY/MM/DD") & _
             IIf(LCase$(status) = LCase$(STATUS_INVOICE_PAID), " (was Paid)", "")
    On Error GoTo 0
    
    ShowInfo "Invoice date updated successfully." & vbNewLine & vbNewLine & _
             "Invoice : " & invNum & vbNewLine & _
             "Customer: " & customer & vbNewLine & _
             "Old date: " & defaultDateStr & vbNewLine & _
             "New date: " & Format(newDate, "YYYY/MM/DD") & vbNewLine & vbNewLine & _
             "Aging will recalculate automatically the next time you generate a statement.", _
             "Date Changed"
    Exit Sub

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    
    On Error Resume Next
    If Not wbTrack Is Nothing Then CloseTrackingWorkbook wbTrack, False
    RestoreExcelState
    LogAudit "ChangeInvoiceDate", invNum, customer, patient, total, _
             "FAILED", "Err#" & errNum & ": " & errDesc
    On Error GoTo 0
    
    ShowError "Error changing invoice date." & vbNewLine & vbNewLine & _
              "Error #" & errNum & ":" & vbNewLine & _
              IIf(Len(errDesc) > 0, errDesc, "(no description provided by Excel)")
End Sub

' =====================================================
' ===== HELPERS =====
' =====================================================

Private Function GetActiveInvoiceNumber() As String
    Dim s As String
    GetActiveInvoiceNumber = ""
    
    On Error Resume Next
    If ActiveSheet Is Nothing Then Exit Function
    
    ' Only auto-detect if we're on the Invoice sheet
    If StrComp(ActiveSheet.Name, SHEET_INVOICE, vbTextCompare) <> 0 Then Exit Function
    
    s = Trim(CStr(ActiveSheet.Range(CELL_DOC_NUMBER).Value))
    GetActiveInvoiceNumber = s
    On Error GoTo 0
End Function

Private Function FindInvoiceRowByNumber(ws As Worksheet, ByVal invNum As String) As Long
    Dim lastRow As Long, i As Long
    Dim candidate As String, target As String
    
    FindInvoiceRowByNumber = 0
    If ws Is Nothing Then Exit Function
    
    target = Trim(CStr(invNum))
    If target = "" Then Exit Function
    
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Function
    
    ' Pass 1: exact case-insensitive match
    For i = 2 To lastRow
        candidate = Trim(CStr(ws.Cells(i, COL_I_NUMBER).Value))
        If candidate <> "" Then
            If StrComp(candidate, target, vbTextCompare) = 0 Then
                FindInvoiceRowByNumber = i
                Exit Function
            End If
        End If
    Next i
    
    ' Pass 2: fallback - match ignoring INVOICE_PREFIX
    For i = 2 To lastRow
        candidate = Trim(CStr(ws.Cells(i, COL_I_NUMBER).Value))
        If candidate <> "" Then
            If StrComp(NormaliseInvNum(candidate), NormaliseInvNum(target), vbTextCompare) = 0 Then
                FindInvoiceRowByNumber = i
                Exit Function
            End If
        End If
    Next i
End Function

' Strips the configured INVOICE_PREFIX and whitespace so "INV-0067",
' "inv-0067" and "0067" all collapse to "0067" for matching purposes.
Private Function NormaliseInvNum(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim(CStr(s)))
    t = Replace(t, UCase$(INVOICE_PREFIX) & "-", "")
    t = Replace(t, UCase$(INVOICE_PREFIX), "")
    t = Replace(t, " ", "")
    NormaliseInvNum = t
End Function

Private Function FormatDateOrBlank(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatDateOrBlank = Format(CDate(v), "YYYY/MM/DD")
    Else
        FormatDateOrBlank = "(none)"
    End If
End Function

' Asks for a new date. Tries PickDate first; falls back to an InputBox
' if PickDate isn't available (which shouldn't happen after Stage 1, but
' the fallback keeps this sub robust against module-removal accidents).
Private Function TryPickDateOrPrompt(ByVal defaultDateStr As String, _
                                     ByVal invNum As String, _
                                     ByVal customer As String, _
                                     ByVal patient As String) As Date
    Dim picked As Date
    Dim newDateStr As String
    Dim defaultDate As Date
    
    TryPickDateOrPrompt = 0
    
    ' Parse the default date string once
    If IsDate(defaultDateStr) Then
        defaultDate = CDate(defaultDateStr)
    Else
        defaultDate = Date
    End If
    
    ' --- Try the proper date picker first ---
    picked = PickDate(defaultDate)
    If picked > 0 Then
        TryPickDateOrPrompt = picked
        Exit Function
    End If
    
    ' --- User cancelled the picker. Offer InputBox fallback only if they
    '     specifically asked for it. Cancelling the picker = cancelling
    '     the change, which matches the rest of the app's behaviour. ---
    Exit Function
End Function

Private Sub TryUpdateInvoiceSheetDate(ByVal newDate As Date)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    If ws Is Nothing Then Exit Sub
    
    ws.Range(CELL_DOC_DATE).Value = newDate
    ws.Range(CELL_DOC_DATE).NumberFormat = "YYYY/MM/DD"
    On Error GoTo 0
End Sub

