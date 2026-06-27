VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDateButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================
' ===== CLASS TO HANDLE DAY BUTTON CLICKS =====
' =====================================================

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Private mForm As frmDatePicker

Public Sub Init(ByRef frm As frmDatePicker, ByRef button As MSForms.CommandButton)
    Set mForm = frm
    Set btn = button
End Sub

Private Sub btn_Click()
    Dim dayNum As Integer
    
    If mForm Is Nothing Then Exit Sub
    If btn Is Nothing Then Exit Sub
    If Not btn.Visible Then Exit Sub
    If Not btn.Enabled Then Exit Sub
    If Not IsNumeric(btn.Tag) Then Exit Sub
    
    dayNum = CInt(btn.Tag)
    
    If dayNum >= 1 And dayNum <= 31 Then
        mForm.DayClicked dayNum
    End If
End Sub
