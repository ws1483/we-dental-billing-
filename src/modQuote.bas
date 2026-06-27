Attribute VB_Name = "modQuote"
Option Explicit

' =====================================================
' ===== QUOTE - NOTES BLOCK CONSTANTS =====
' =====================================================
' Rows 38-39, columns A-H, merged as one big notes box.

Private Const QUOTE_NOTES_TOP_ROW    As Long = 38
Private Const QUOTE_NOTES_BOTTOM_ROW As Long = 39
Private Const QUOTE_NOTES_RANGE      As String = "A38:H39"

' =====================================================
' ===== GET NEXT QUOTE NUMBER =====
' =====================================================
' Reads next number from tracker but does NOT persist
' the Settings counter - that happens after a successful export.

Public Function GetNextQuoteNumber() As String
    Dim nextQ As String
    
    nextQ = GetNextQuoteNumberFromTracking()
    
    If nextQ <> "ERROR" Then
        GetNextQuoteNumber = nextQ
        Exit Function
    End If
    
    ' Fallback: use Settings counter if tracking file unavailable
    Dim lastNum As Long
    
    EnsureSettingsSheet
    If Not SheetExists(SHEET_SETTINGS) Then
        ShowError "Settings sheet not found."
        GetNextQuoteNumber = "ERROR"
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    lastNum = ThisWorkbook.Sheets(SHEET_SETTINGS).Range("B1").Value
    GetNextQuoteNumber = QUOTE_PREFIX & "-" & (lastNum + 1)
    Exit Function

ErrorHandler:
    ShowError "Error generating quote number: " & Err.Description
    GetNextQuoteNumber = "ERROR"
End Function

' =====================================================
' ===== RESET QUOTE NOTES BLOCK =====
' =====================================================
' Wipes anything the user typed into the notes box, re-merges
' it, and drops the bold "Note:" label back in.

Public Sub ResetQuoteNotesBlock(ws As Worksheet)
    On Error Resume Next
    
    With ws.Range(QUOTE_NOTES_RANGE)
        .UnMerge
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.bold = False
        .Font.Italic = False
        .Font.Size = 11
        .Font.Color = vbBlack
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
    ' Re-merge rows 38 + 39 across columns A-H as one big notes box
    ws.Range(QUOTE_NOTES_RANGE).Merge
    
    ' Put the bold "Note:" label back into the top-left of the merged cell
    With ws.Range("A" & QUOTE_NOTES_TOP_ROW)
        .Value = "Note:"
        .Font.bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
    ' Match the row heights so the box has room to type
    ws.rows(QUOTE_NOTES_TOP_ROW).rowHeight = 18
    ws.rows(QUOTE_NOTES_BOTTOM_ROW).rowHeight = 18
    
    On Error GoTo 0
End Sub

' =====================================================
' ===== CLEAR AND CREATE NEW QUOTE =====
' =====================================================

Public Sub ClearAndNewQuote()
    Dim ws As Worksheet
    Dim i As Integer
    Dim newQuoteNum As String
    
    If Not SheetExists(SHEET_QUOTE) Then
        ShowError "Quote sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_QUOTE)
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Creating new quote..."
    
    ' Clear customer info
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
    
    ' Clear line items
    For i = START_ROW To END_ROW
        On Error Resume Next
        ws.Cells(i, 1).ClearContents
        ws.Cells(i, 4).ClearContents
        On Error GoTo ErrorHandler
    Next i
    
    RestoreQuoteFormulas
    
    ' Reset the merged Notes block on rows 38-39
    ResetQuoteNotesBlock ws
    
    newQuoteNum = GetNextQuoteNumber()
    If newQuoteNum = "ERROR" Then GoTo CleanUp
    
    ws.Range(CELL_DOC_DATE).Value = Date
    ws.Range(CELL_DOC_DATE).NumberFormat = "YYYY/MM/DD"
    ws.Range(CELL_DOC_NUMBER).Value = newQuoteNum
    
    ws.Activate
    ws.Range(CELL_CUSTOMER).Select
    
    RestoreExcelState
    ShowInfo "New Quote Created: " & newQuoteNum, "New Quote"
    Exit Sub

CleanUp:
    RestoreExcelState
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error creating new quote: " & Err.Description
End Sub

' =====================================================
' ===== EXPORT QUOTE TO PDF =====
' =====================================================

Public Sub ExportQuoteToPDF()
    Dim ws As Worksheet
    Dim savePath As String, fileName As String
    Dim patientName As String, quoteNum As String
    
    If Not SheetExists(SHEET_QUOTE) Then
        ShowError "Quote sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_QUOTE)
    
    If Trim(ws.Range(CELL_DOC_NUMBER).Value) = "" Then
        ShowWarning "Please generate a quote number first.", "Missing Quote Number"
        Exit Sub
    End If
    
    patientName = CleanFileName(SafeString(ws.Range(CELL_PATIENT).Value))
    If patientName = "" Then patientName = CleanFileName(SafeString(ws.Range(CELL_CUSTOMER).Value))
    If patientName = "" Then patientName = "Quote"
    quoteNum = CleanFileName(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    fileName = patientName & " - " & quoteNum & ".pdf"
    
    savePath = GetSavePath(fileName, "PDF Files (*.pdf), *.pdf", "Save Quote as PDF")
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    SafeStatusBar "Exporting PDF..."
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=savePath, Quality:=xlQualityStandard
    
    SafeStatusBar False
    ShowInfo "PDF exported to:" & vbNewLine & vbNewLine & savePath, "Export Complete"
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error exporting PDF: " & Err.Description
End Sub

' =====================================================
' ===== EXPORT QUOTE TO EXCEL (FULL WORKBOOK COPY) =====
' =====================================================

Public Sub ExportQuoteToExcel()
    Dim ws As Worksheet
    Dim savePath As String
    Dim fileName As String
    Dim patientName As String, quoteNum As String
    Dim customer As String, patient As String
    Dim vSubTotal As Variant, vDiscount As Variant, vVAT As Variant, vTotal As Variant
    Dim saveOk As Boolean
    
    If Not SheetExists(SHEET_QUOTE) Then
        ShowError "Quote sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_QUOTE)
    
    If Trim(ws.Range(CELL_DOC_NUMBER).Value) = "" Then
        ShowWarning "Please generate a quote number first.", "Missing Quote Number"
        Exit Sub
    End If
    
    patientName = CleanFileName(SafeString(ws.Range(CELL_PATIENT).Value))
    If patientName = "" Then patientName = CleanFileName(SafeString(ws.Range(CELL_CUSTOMER).Value))
    If patientName = "" Then patientName = "Quote"
    quoteNum = CleanFileName(SafeString(ws.Range(CELL_DOC_NUMBER).Value))
    fileName = patientName & " - " & quoteNum & ".xlsm"
    
    savePath = GetSavePath(fileName, "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", "Save Quote as Excel")
    If savePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Exporting quote (full workbook)..."
    
    customer = SafeString(ws.Range(CELL_CUSTOMER).Value)
    patient = SafeString(ws.Range(CELL_PATIENT).Value)
    
    vSubTotal = ws.Range(CELL_SUMMARY_SUBTOT).Value
    vDiscount = ws.Range(CELL_SUMMARY_DISC).Value
    vVAT = ws.Range(CELL_SUMMARY_VAT).Value
    vTotal = ws.Range(CELL_SUMMARY_TOTAL).Value
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs savePath
    Application.DisplayAlerts = True
    
    ' ---- Save AND verify ----
    saveOk = SaveAndVerifyQuote(SafeString(ws.Range(CELL_DOC_NUMBER).Value), _
                                 SafeDate(ws.Range(CELL_DOC_DATE).Value), _
                                 customer, patient, _
                                 SafeString(ws.Range(CELL_APPLIANCE).Value), _
                                 vSubTotal, vDiscount, vVAT, vTotal, _
                                 STATUS_QUOTE_OPEN, "", savePath)
    
    PersistQuoteCounter SafeString(ws.Range(CELL_DOC_NUMBER).Value)
    RestoreExcelState
    
    If saveOk Then
        ShowVerifiedConfirmation "Quote", quoteNum, customer, patient, _
                                 r2(vTotal), _
                                 "Exported to: " & savePath
    Else
        ShowVerifyFailure "Quote", quoteNum, customer, patient, r2(vTotal)
    End If
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error exporting quote: " & Err.Description
End Sub

' =====================================================
' ===== PRINT QUOTE =====
' =====================================================

Public Sub PrintQuote()
    If Not SheetExists(SHEET_QUOTE) Then
        ShowError "Quote sheet not found."
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    ThisWorkbook.Sheets(SHEET_QUOTE).PrintOut Copies:=1, Preview:=True
    Exit Sub

ErrorHandler:
    ShowError "Error printing quote: " & Err.Description
End Sub

