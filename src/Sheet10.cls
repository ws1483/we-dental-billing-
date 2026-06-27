VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =====================================================
' ===== MENU SHEET CODE-BEHIND =====
' Double-click handler for:
'   - Opening linked Quote / Invoice files (column C)
'   - Pagination on Open Quotes panel (Prev/Next cells)
'   - Pagination on Unpaid Invoices panel (Prev/Next cells)
'   - Auto-refresh when customer dropdown changes (Worksheet_Change)
'
' Uses BeforeDoubleClick (not SelectionChange) so:
'   - Single clicks never accidentally trigger actions
'   - Cancel = True suppresses Excel's in-cell edit mode
' =====================================================

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    Dim r As Long
    Dim docNum As String
    Dim addr As String
    
    ' Only respond to single-cell double-clicks
    If target.Cells.count <> 1 Then Exit Sub
    
    addr = target.Address(False, False)   ' e.g. "F28"
    r = target.row
    
    ' ============================================================
    ' === Pagination buttons (check these BEFORE column C checks) ===
    ' ============================================================
    
    ' --- Open Quotes: Prev ---
    If StrComp(addr, MENU_OPEN_QUOTES_PREV_CELL, vbTextCompare) = 0 Then
        Cancel = True
        OpenQuotesPrevPage
        Exit Sub
    End If
    
    ' --- Open Quotes: Next ---
    If StrComp(addr, MENU_OPEN_QUOTES_NEXT_CELL, vbTextCompare) = 0 Then
        Cancel = True
        OpenQuotesNextPage
        Exit Sub
    End If
    
    ' --- Unpaid Invoices: Prev ---
    If StrComp(addr, MENU_UNPAID_INV_PREV_CELL, vbTextCompare) = 0 Then
        Cancel = True
        UnpaidInvoicesPrevPage
        Exit Sub
    End If
    
    ' --- Unpaid Invoices: Next ---
    If StrComp(addr, MENU_UNPAID_INV_NEXT_CELL, vbTextCompare) = 0 Then
        Cancel = True
        UnpaidInvoicesNextPage
        Exit Sub
    End If
    
    ' ============================================================
    ' === Document open (column C in panel data rows) ===
    ' ============================================================
    
    If target.Column <> 3 Then Exit Sub
    
    ' --- Open Quotes panel range ---
    If r >= MENU_OPEN_QUOTES_FIRST_DATA And r <= MENU_OPEN_QUOTES_LAST_DATA Then
        docNum = Trim(CStr(Me.Cells(r, "C").Value))
        If docNum <> "" Then
            Cancel = True
            OpenLinkedFileForRow r, "Quote"
        End If
        Exit Sub
    End If
    
    ' --- Unpaid Invoices panel range ---
    If r >= MENU_UNPAID_INV_FIRST_DATA And _
       r <= MENU_UNPAID_INV_FIRST_DATA + MENU_UNPAID_INV_MAX_ROWS - 1 Then
        docNum = Trim(CStr(Me.Cells(r, "C").Value))
        If docNum <> "" Then
            Cancel = True
            OpenLinkedFileForRow r, "Invoice"
        End If
    End If
End Sub

' =====================================================
' ===== AUTO-REFRESH ON FILTER CHANGE =====
' =====================================================
' When the user picks a different customer in the dropdown, reset
' the page index to 1 and refresh the panel automatically.
' (Without this, you'd see "Page 3 of 1" briefly after switching filters.)

Private Sub Worksheet_Change(ByVal target As Range)
    Dim addr As String
    
    If target.Cells.count <> 1 Then Exit Sub
    addr = target.Address(False, False)
    
    Application.EnableEvents = False
    On Error Resume Next
    
    Select Case addr
        Case MENU_OPEN_QUOTES_CUST_CELL
            OpenQuotesResetToFirstPage
            RefreshOpenQuotesList
            
        Case MENU_UNPAID_INV_CUST_CELL
            UnpaidInvoicesResetToFirstPage
            RefreshUnpaidInvoicesList
    End Select
    
    On Error GoTo 0
    Application.EnableEvents = True
End Sub

' =====================================================
' ===== OPEN LINKED FILE FROM PANEL ROW =====
' =====================================================
' Reads the file path from the hidden column I on the row
' and opens it with the registered default application
' (via ShellExecute, not the deprecated rundll32 hack).

Private Sub OpenLinkedFileForRow(ByVal rowNumber As Long, ByVal docKind As String)
    Dim filePath As String
    Dim docNum As String
    
    docNum = Trim(CStr(Me.Cells(rowNumber, "C").Value))
    filePath = Trim(CStr(Me.Cells(rowNumber, "I").Value))
    
    If filePath = "" Then
        ShowWarning "No file path stored for " & LCase$(docKind) & " " & docNum & "." & vbNewLine & vbNewLine & _
                    "Re-export the " & LCase$(docKind) & " to record its file path in the tracker.", _
                    "No File Linked"
        Exit Sub
    End If
    
    If Not FileExists(filePath) Then
        ShowError docKind & " file not found:" & vbNewLine & vbNewLine & filePath & vbNewLine & vbNewLine & _
                  "The file may have been moved, renamed or deleted." & vbNewLine & _
                  "Re-export the " & LCase$(docKind) & " to refresh the link."
        Exit Sub
    End If
    
    OpenWithDefaultApp filePath
End Sub

