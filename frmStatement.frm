VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStatement 
   Caption         =   "Generate Statement"
   ClientHeight    =   4440
   ClientLeft      =   12
   ClientTop       =   84
   ClientWidth     =   7380
   OleObjectBlob   =   "frmStatement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================
' ===== STATEMENT FORM CODE-BEHIND =====
' Uses these controls:
'   cboCust    - Customer ComboBox
'   txtFrom    - From Date TextBox
'   txtTo      - To Date TextBox
'   btnPkFrom  - From Date picker button
'   btnPkTo    - To Date picker button
'   btnGen     - Generate Statement button
'   btnCnl     - Cancel button
' =====================================================

Private Sub UserForm_Initialize()
    Me.caption = "Generate Statement"
    
    ' Defaults: month-to-date
    Me.txtFrom.Value = Format(DateSerial(Year(Date), Month(Date), 1), "YYYY/MM/DD")
    Me.txtTo.Value = Format(Date, "YYYY/MM/DD")
    
    ' Populate customer combo
    Dim col As Collection, item As Variant
    Me.cboCust.Clear
    Set col = GetCustomerList()
    If Not col Is Nothing Then
        For Each item In col
            Me.cboCust.AddItem CStr(item)
        Next item
    End If
    
    ' --- Combo behaviour tweaks so the dropdown is properly scrollable ---
    With Me.cboCust
        .style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryNone
        .ListRows = 12
        .ColumnCount = 1
        .BoundColumn = 1
        .TabStop = True
    End With
End Sub

' --- Keyboard scroll fallback (mouse wheel inside a VBA dropdown is unreliable) ---
Private Sub cboCust_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyPageDown
            If Me.cboCust.ListCount = 0 Then Exit Sub
            If Me.cboCust.ListIndex < Me.cboCust.ListCount - 5 Then
                Me.cboCust.ListIndex = Me.cboCust.ListIndex + 5
            Else
                Me.cboCust.ListIndex = Me.cboCust.ListCount - 1
            End If
        Case vbKeyPageUp
            If Me.cboCust.ListCount = 0 Then Exit Sub
            If Me.cboCust.ListIndex > 5 Then
                Me.cboCust.ListIndex = Me.cboCust.ListIndex - 5
            Else
                Me.cboCust.ListIndex = 0
            End If
    End Select
End Sub

' --- From-Date picker button ---
Private Sub btnPkFrom_Click()
    Dim picked As Date
    Dim defaultDate As Date
    
    If IsDate(Me.txtFrom.Value) Then
        defaultDate = CDate(Me.txtFrom.Value)
    Else
        defaultDate = Date
    End If
    
    picked = PickDate(defaultDate)
    If picked > 0 Then
        Me.txtFrom.Value = Format(picked, "YYYY/MM/DD")
    End If
End Sub

' --- To-Date picker button ---
Private Sub btnPkTo_Click()
    Dim picked As Date
    Dim defaultDate As Date
    
    If IsDate(Me.txtTo.Value) Then
        defaultDate = CDate(Me.txtTo.Value)
    Else
        defaultDate = Date
    End If
    
    picked = PickDate(defaultDate)
    If picked > 0 Then
        Me.txtTo.Value = Format(picked, "YYYY/MM/DD")
    End If
End Sub

' --- Generate Statement ---
' Validate everything FIRST, then unload the form, THEN run generation.
' This avoids the "Me is gone but procedure still references it" trap.
Private Sub btnGen_Click()
    Dim custName As String
    Dim fromDate As Date, toDate As Date
    
    custName = Trim(CStr(Me.cboCust.Value))
    If custName = "" Then
        MsgBox "Please select a customer.", vbExclamation, "Missing Customer"
        Me.cboCust.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.txtFrom.Value) Then
        MsgBox "Please enter a valid From Date (YYYY/MM/DD).", vbExclamation, "Invalid Date"
        Me.txtFrom.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.txtTo.Value) Then
        MsgBox "Please enter a valid To Date (YYYY/MM/DD).", vbExclamation, "Invalid Date"
        Me.txtTo.SetFocus
        Exit Sub
    End If
    
    fromDate = CDate(Me.txtFrom.Value)
    toDate = CDate(Me.txtTo.Value)
    
    If fromDate > toDate Then
        MsgBox "From Date cannot be after To Date.", vbExclamation, "Invalid Range"
        Me.txtFrom.SetFocus
        Exit Sub
    End If
    
    ' --- All validation passed. Release the form, then generate. ---
    Unload Me
    GenerateStatementWithParams custName, fromDate, toDate
End Sub

' --- Cancel ---
Private Sub btnCnl_Click()
    Unload Me
End Sub
