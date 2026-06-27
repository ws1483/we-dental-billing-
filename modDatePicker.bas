Attribute VB_Name = "modDatePicker"
Option Explicit

' =====================================================
' ===== REUSABLE DATE PICKER FUNCTION =====
' =====================================================
' Returns the selected date, or 0 if cancelled.
'
' Implementation note:
' Uses a module-level singleton form instance so that the
' frmDatePicker.UserForm_Initialize "build 42 buttons once"
' optimisation actually pays off across multiple calls.
' The instance is kept alive until the workbook closes
' (see Workbook_BeforeClose -> ReleaseDatePicker).

Private mPicker As frmDatePicker

Public Function PickDate(Optional ByVal defaultDate As Date = 0) As Date
    ' Lazy-create the singleton on first use
    If mPicker Is Nothing Then
        Set mPicker = New frmDatePicker
    End If
    
    If defaultDate > 0 Then
        mPicker.SetInitialDate defaultDate
    End If
    
    ' Reset cancelled flag explicitly in case the form is being reused
    mPicker.Cancelled = True
    
    mPicker.Show vbModal
    
    If mPicker.Cancelled Then
        PickDate = 0
    Else
        PickDate = mPicker.SelectedDate
    End If
    
    ' Note: do NOT unload — keep the instance alive for the next call.
End Function

' Call from Workbook_BeforeClose to release the singleton cleanly.
Public Sub ReleaseDatePicker()
    On Error Resume Next
    If Not mPicker Is Nothing Then
        Unload mPicker
        Set mPicker = Nothing
    End If
    On Error GoTo 0
End Sub
