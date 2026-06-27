Attribute VB_Name = "modForms"
Option Explicit

' =====================================================
' ===== SHOW MAIN MENU =====
' =====================================================
' Navigates to the Menu SHEET (the old frmMainMenu UserForm
' has been replaced by a sheet-based dashboard).

Public Sub ShowMainMenu()
    On Error Resume Next
    
    If Not SheetExists(SHEET_MENU) Then
        ' Menu sheet missing — build it
                MsgBox "The '" & SHEET_MENU & "' sheet is missing from this workbook." & vbNewLine & vbNewLine & _
               "Please restore from a backup or contact the developer.", _
               vbCritical, "Menu Sheet Missing"
        Exit Sub
    End If
    
    HideBackgroundSheets
    HideWorkingSheets
    
    ThisWorkbook.Sheets(SHEET_MENU).Visible = xlSheetVisible
    ThisWorkbook.Sheets(SHEET_MENU).Activate
    ThisWorkbook.Sheets(SHEET_MENU).Range("A1").Select
    
    RefreshOpenQuotesList
    On Error GoTo 0
End Sub

' =====================================================
' ===== QUICK ACCESS SUBS (for sheet buttons) =====
' =====================================================
' These keep existing buttons on Quote / Invoice / Statement
' sheets working. Each one just delegates to the real
' implementation in modQuote / modInvoice / modStatement / etc.

Public Sub BtnNewQuote()
    ClearAndNewQuote
End Sub

Public Sub BtnExportQuotePDF()
    ExportQuoteToPDF
End Sub

Public Sub BtnExportQuoteExcel()
    ExportQuoteToExcel
End Sub

Public Sub BtnSyncQuote()
    SyncQuoteFromFile
End Sub

Public Sub BtnNewInvoice()
    CreateNewInvoice
End Sub

Public Sub BtnConvertToInvoice()
    ConvertToInvoice
End Sub

Public Sub BtnExportInvoicePDF()
    ExportInvoiceToPDF
End Sub

Public Sub BtnExportInvoiceExcel()
    ExportInvoiceToExcel
End Sub

Public Sub BtnSyncInvoice()
    SyncInvoiceFromFile
End Sub

Public Sub btnMarkPaid()
    ShowMarkAsPaidForm
End Sub

Public Sub BtnGenerateStatement()
    ShowStatementForm
End Sub

Public Sub BtnExportStatementPDF()
    ExportStatementToPDF
End Sub

Public Sub BtnExportStatementExcel()
    ExportStatementToExcel
End Sub

Public Sub BtnPrintStatement()
    PrintStatement
End Sub

Public Sub BtnRestoreQuoteFormulas()
    RestoreQuoteFormulas
    ShowInfo "Quote formulas restored.", "Complete"
End Sub

Public Sub BtnRestoreInvoiceFormulas()
    RestoreInvoiceFormulas
    ShowInfo "Invoice formulas restored.", "Complete"
End Sub

Public Sub BtnRestoreAllFormulas()
    RestoreQuoteFormulas
    RestoreInvoiceFormulas
    ShowInfo "All formulas restored.", "Complete"
End Sub

Public Sub BtnAddDropdowns()
    AddDescriptionDropdowns
End Sub

Public Sub BtnSetupAll()
    SetupAllFormulasAndDropdowns
End Sub

Public Sub BtnPrintQuote()
    PrintQuote
End Sub

Public Sub BtnPrintInvoice()
    PrintInvoice
End Sub

' =====================================================
' ===== DATE PICKER FOR INVOICE SHEET =====
' =====================================================
' Assign this macro to the [...] button next to the date cell on the Invoice sheet.
' If the user cancels the picker, the existing date is preserved (not silently
' overwritten with today's date).

Public Sub BtnPickInvoiceDate()
    Dim ws As Worksheet
    Dim pickedDate As Date
    Dim currentDate As Date
    
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    
    If IsDate(ws.Range(CELL_DOC_DATE).Value) Then
        currentDate = CDate(ws.Range(CELL_DOC_DATE).Value)
    Else
        currentDate = Date
    End If
    
    pickedDate = PickDate(currentDate)
    
    ' Only overwrite if the user actually picked a date.
    ' PickDate returns 0 on Cancel.
    If pickedDate > 0 Then
        ws.Range(CELL_DOC_DATE).Value = pickedDate
        ws.Range(CELL_DOC_DATE).NumberFormat = "YYYY/MM/DD"
    End If
End Sub

' =====================================================
' ===== UPDATE TO TRACKER BUTTONS =====
' =====================================================

' Assign this to a button on the QUOTE sheet
Public Sub BtnUpdateQuoteToTracker()
    UpdateQuoteToTracker
End Sub

' Assign this to a button on the INVOICE sheet
Public Sub BtnUpdateInvoiceToTracker()
    UpdateInvoiceToTracker
End Sub
