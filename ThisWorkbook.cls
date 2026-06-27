VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =====================================================
' ===== WORKBOOK EVENTS =====
' =====================================================

Private Sub Workbook_Open()
    Dim wsMenu As Worksheet
    
    On Error GoTo ErrorHandler
    
    If Not SheetExists(SHEET_MENU) Then
        MsgBox "The '" & SHEET_MENU & "' sheet is missing from this workbook." & vbNewLine & vbNewLine & _
               "This is required for the app to start. Please restore from a backup, " & _
               "or contact the developer if you don't have one.", _
               vbCritical, "Menu Sheet Missing"
        Exit Sub
    End If
    
    Set wsMenu = ThisWorkbook.Sheets(SHEET_MENU)
    
    ' Hide working + background sheets so only the menu is visible
    HideBackgroundSheets
    HideWorkingSheets
    
    wsMenu.Visible = xlSheetVisible
    wsMenu.Activate
    
    ' Refresh both panels at startup so the user sees current data
    On Error Resume Next
    RefreshOpenQuotesList
    RefreshUnpaidInvoicesList
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If SheetExists(SHEET_MENU) Then
        ThisWorkbook.Sheets(SHEET_MENU).Visible = xlSheetVisible
        ThisWorkbook.Sheets(SHEET_MENU).Activate
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    
    On Error GoTo 0
End Sub

