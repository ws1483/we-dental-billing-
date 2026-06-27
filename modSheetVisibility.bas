Attribute VB_Name = "modSheetVisibility"
Option Explicit

' =====================================================
' ===== SHEET VISIBILITY HELPERS =====
' =====================================================
' Lives in a regular module (not ThisWorkbook) so callers in
' other modules can see these Public subs by their bare names.

Public Sub HideBackgroundSheets()
    HidePipeSeparatedSheets MENU_BACKGROUND_SHEETS, xlSheetVeryHidden
End Sub

Public Sub HideWorkingSheets()
    HidePipeSeparatedSheets MENU_WORKING_SHEETS, xlSheetHidden
End Sub

Public Sub ShowAllWorkingSheets()
    ShowPipeSeparatedSheets MENU_WORKING_SHEETS
End Sub

' =====================================================
' ===== INTERNAL =====
' =====================================================

Private Sub HidePipeSeparatedSheets(ByVal pipeList As String, _
                                     ByVal visibility As XlSheetVisibility)
    Dim names() As String
    Dim i As Long
    Dim ws As Worksheet
    
    If Trim$(pipeList) = "" Then Exit Sub
    
    names = Split(pipeList, "|")
    
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(Trim$(names(i)))
        If Not ws Is Nothing Then
            ws.Visible = visibility
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub ShowPipeSeparatedSheets(ByVal pipeList As String)
    Dim names() As String
    Dim i As Long
    Dim ws As Worksheet
    
    If Trim$(pipeList) = "" Then Exit Sub
    
    names = Split(pipeList, "|")
    
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(Trim$(names(i)))
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
        End If
        On Error GoTo 0
    Next i
End Sub

