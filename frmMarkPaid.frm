VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMarkPaid 
   Caption         =   "Mark Invoice as Paid"
   ClientHeight    =   9840.001
   ClientLeft      =   312
   ClientTop       =   360
   ClientWidth     =   11580
   OleObjectBlob   =   "frmMarkPaid.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMarkPaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================
' ===== MARK INVOICE PAID FORM =====
' Controls on the form (verified names):
'   cboCustomer       - Customer filter dropdown
'   btnRefresh        - Reload invoice list
'   lstInvoices       - Multi-select ListBox of unpaid/partial invoices
'   chkSelectAll      - Toggles all rows in lstInvoices on/off
'   btnPkDate         - Opens date picker for payment date
'   txtPaymentDate    - Displays selected payment date (YYYY/MM/DD)
'   btnMarkPaid       - Marks ALL ticked invoices as fully Paid on the chosen date
'   btnApplyPayment   - Asks for an amount, applies to SINGLE selected invoice
'   btnCancel         - Close form
' =====================================================

Private mIsLoading As Boolean

Private Sub UserForm_Initialize()
    mIsLoading = True
    
    Me.caption = "Mark Invoice as Paid"
    
    ' Default payment date = today
    Me.txtPaymentDate.Value = Format(Date, "YYYY/MM/DD")
    
    ' Multi-select listbox so user can tick several invoices
    With Me.lstInvoices
        .ColumnCount = 5
        .ColumnWidths = "70;75;110;90;55"   ' Date | Invoice# | Customer | Patient | Balance
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption      ' shows checkboxes next to each row
    End With
    
    Me.chkSelectAll.Value = False
    
    LoadCustomerFilter
    LoadUnpaidInvoices
    
    mIsLoading = False
    
    UpdateButtonStates
End Sub

' --- Customer dropdown populates from Customers sheet ---
Private Sub LoadCustomerFilter()
    Dim col As Collection, item As Variant
    
    Me.cboCustomer.Clear
    Me.cboCustomer.AddItem "(All Customers)"
    
    Set col = GetCustomerList()
    If Not col Is Nothing Then
        For Each item In col
            Me.cboCustomer.AddItem CStr(item)
        Next item
    End If
    
    With Me.cboCustomer
        .style = fmStyleDropDownList
        .MatchEntry = fmMatchEntryNone
        .ListRows = 12
        .ListIndex = 0
    End With
End Sub

' --- Customer change reloads list ---
Private Sub cboCustomer_Change()
    If mIsLoading Then Exit Sub
    LoadUnpaidInvoices
    UpdateButtonStates
End Sub

' --- Refresh button ---
Private Sub btnRefresh_Click()
    LoadUnpaidInvoices
    Me.chkSelectAll.Value = False
    UpdateButtonStates
End Sub

' --- Populate the listbox with all unpaid/partial invoices ---
Private Sub LoadUnpaidInvoices()
    Dim wbTrack As Workbook, wsTrack As Worksheet
    Dim trackingPath As String
    Dim lastRow As Long, r As Long
    Dim invNum As String, cust As String, patient As String
    Dim status As String, total As Double, paid As Double, credit As Double, balance As Double
    Dim invDate As Variant
    Dim selectedCust As String
    Dim filterToCust As Boolean
    
    mIsLoading = True
    Me.lstInvoices.Clear
    
    selectedCust = Trim(CStr(Me.cboCustomer.Value))
    filterToCust = (selectedCust <> "" And selectedCust <> "(All Customers)")
    
    trackingPath = GetInvoiceTrackingPath()
    If trackingPath = "" Then
        mIsLoading = False
        Exit Sub
    End If
    
    Set wbTrack = OpenTrackingWorkbook(trackingPath, True)
    If wbTrack Is Nothing Then
        mIsLoading = False
        Exit Sub
    End If
    
    On Error GoTo CleanUp
    
    If Not SheetExistsInWorkbook(TRACK_SHEET_INVOICES, wbTrack) Then GoTo CleanUp
    
    Set wsTrack = wbTrack.Sheets(TRACK_SHEET_INVOICES)
    lastRow = wsTrack.Cells(wsTrack.rows.count, "A").End(xlUp).row
    
    For r = 2 To lastRow
        invNum = Trim(CStr(wsTrack.Cells(r, COL_I_NUMBER).Value))
        status = Trim(CStr(wsTrack.Cells(r, COL_I_STATUS).Value))
        
        If invNum <> "" Then
            If LCase$(status) <> LCase$(STATUS_INVOICE_PAID) And _
               LCase$(status) <> LCase$(STATUS_INVOICE_CANCELLED) Then
                
                cust = Trim(CStr(wsTrack.Cells(r, COL_I_CUSTOMER).Value))
                
                If (Not filterToCust) Or (StrComp(cust, selectedCust, vbTextCompare) = 0) Then
                    invDate = wsTrack.Cells(r, COL_I_DATE).Value
                    patient = Trim(CStr(wsTrack.Cells(r, COL_I_PATIENT).Value))
                    total = r2(wsTrack.Cells(r, COL_I_TOTAL).Value)
                    paid = r2(wsTrack.Cells(r, COL_I_PAIDAMOUNT).Value)
                    credit = r2(wsTrack.Cells(r, COL_I_CREDITAMT).Value)
                    
                    balance = Round(total - paid - credit, 2)
                    If balance > 0 Then
                        Me.lstInvoices.AddItem
                        Dim newRow As Long
                        newRow = Me.lstInvoices.ListCount - 1
                        
                        If IsDate(invDate) Then
                            Me.lstInvoices.List(newRow, 0) = Format(CDate(invDate), "YYYY/MM/DD")
                        Else
                            Me.lstInvoices.List(newRow, 0) = ""
                        End If
                        Me.lstInvoices.List(newRow, 1) = invNum
                        Me.lstInvoices.List(newRow, 2) = cust
                        Me.lstInvoices.List(newRow, 3) = patient
                        Me.lstInvoices.List(newRow, 4) = FmtR(balance)
                    End If
                End If
            End If
        End If
    Next r
    
CleanUp:
    On Error Resume Next
    CloseTrackingWorkbook wbTrack, False
    On Error GoTo 0
    
    mIsLoading = False
End Sub

' --- Select All checkbox toggles every row in the listbox ---
Private Sub chkSelectAll_Click()
    If mIsLoading Then Exit Sub
    
    Dim i As Long
    Dim newState As Boolean
    newState = (Me.chkSelectAll.Value = True)
    
    mIsLoading = True
    For i = 0 To Me.lstInvoices.ListCount - 1
        Me.lstInvoices.Selected(i) = newState
    Next i
    mIsLoading = False
    
    UpdateButtonStates
End Sub

' --- Selection changes update button enablement ---
Private Sub lstInvoices_Change()
    If mIsLoading Then Exit Sub
    UpdateButtonStates
End Sub

Private Sub UpdateButtonStates()
    Dim count As Long
    count = CountSelectedInvoices()
    
    ' Mark Selected as Paid: any selection counts
    Me.btnMarkPaid.Enabled = (count >= 1)
    
    ' Apply Payment (partial): exactly one invoice must be selected
    Me.btnApplyPayment.Enabled = (count = 1)
End Sub

Private Function CountSelectedInvoices() As Long
    Dim i As Long, n As Long
    n = 0
    For i = 0 To Me.lstInvoices.ListCount - 1
        If Me.lstInvoices.Selected(i) Then n = n + 1
    Next i
    CountSelectedInvoices = n
End Function

' --- Date picker button ---
Private Sub btnPkDate_Click()
    Dim picked As Date
    Dim defaultDate As Date
    
    If IsDate(Me.txtPaymentDate.Value) Then
        defaultDate = CDate(Me.txtPaymentDate.Value)
    Else
        defaultDate = Date
    End If
    
    picked = PickDate(defaultDate)
    If picked > 0 Then
        Me.txtPaymentDate.Value = Format(picked, "YYYY/MM/DD")
    End If
End Sub

' --- Mark ALL selected invoices as fully paid on the chosen date ---
Private Sub btnMarkPaid_Click()
    Dim i As Long
    Dim invNum As String
    Dim paymentDate As Date
    Dim count As Long, successCount As Long
    Dim total As Double, paid As Double, credit As Double, due As Double, payAmount As Double
    Dim summary As String
    
    count = CountSelectedInvoices()
    If count < 1 Then
        MsgBox "Please select at least one invoice.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    If Not IsDate(Me.txtPaymentDate.Value) Then
        MsgBox "Please choose a valid Payment Date.", vbExclamation, "Invalid Date"
        Me.btnPkDate.SetFocus
        Exit Sub
    End If
    paymentDate = CDate(Me.txtPaymentDate.Value)
    
    ' Build a summary of what's about to happen
    summary = "Mark " & count & " invoice(s) as fully PAID on " & Format(paymentDate, "YYYY/MM/DD") & "?" & vbNewLine & vbNewLine
    
    For i = 0 To Me.lstInvoices.ListCount - 1
        If Me.lstInvoices.Selected(i) Then
            summary = summary & "  " & CStr(Me.lstInvoices.List(i, 1)) & "   " & _
                      CStr(Me.lstInvoices.List(i, 2)) & "   " & _
                      CStr(Me.lstInvoices.List(i, 4)) & vbNewLine
        End If
    Next i
    
    If MsgBox(summary, vbQuestion + vbYesNo, "Confirm Bulk Payment") <> vbYes Then Exit Sub
    
    successCount = 0
    
    For i = 0 To Me.lstInvoices.ListCount - 1
        If Me.lstInvoices.Selected(i) Then
            invNum = CStr(Me.lstInvoices.List(i, 1))
            
            ReadInvoiceTotals invNum, total, paid, credit
            due = Round(total - credit, 2)
            If due < 0 Then due = 0
            payAmount = Round(due - paid, 2)
            If payAmount < 0 Then payAmount = 0
            
            If payAmount > 0 Then
                ApplyPartialPayment invNum, payAmount, paymentDate
            Else
                UpdateInvoiceStatus invNum, STATUS_INVOICE_PAID, paymentDate
            End If
            
            successCount = successCount + 1
        End If
    Next i
    
    ShowInfo successCount & " invoice(s) marked as Paid on " & Format(paymentDate, "YYYY/MM/DD") & ".", _
             "Payments Recorded"
    
    Me.chkSelectAll.Value = False
    LoadUnpaidInvoices
    UpdateButtonStates
End Sub

' --- Apply a partial payment to a SINGLE selected invoice ---
'     Prompts for the amount because there's no amount textbox on the form.
Private Sub btnApplyPayment_Click()
    Dim i As Long, sel As Long
    Dim invNum As String, custName As String, balanceText As String
    Dim paymentDate As Date
    Dim amtText As String
    Dim payAmount As Double
    Dim total As Double, paid As Double, credit As Double, due As Double, remaining As Double
    
    ' Find the single selected row
    sel = -1
    For i = 0 To Me.lstInvoices.ListCount - 1
        If Me.lstInvoices.Selected(i) Then
            If sel = -1 Then
                sel = i
            Else
                MsgBox "Select ONE invoice for a partial payment.", vbExclamation, "Multiple Selected"
                Exit Sub
            End If
        End If
    Next i
    
    If sel = -1 Then
        MsgBox "Please select one invoice.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    If Not IsDate(Me.txtPaymentDate.Value) Then
        MsgBox "Please choose a valid Payment Date.", vbExclamation, "Invalid Date"
        Me.btnPkDate.SetFocus
        Exit Sub
    End If
    paymentDate = CDate(Me.txtPaymentDate.Value)
    
    invNum = CStr(Me.lstInvoices.List(sel, 1))
    custName = CStr(Me.lstInvoices.List(sel, 2))
    balanceText = CStr(Me.lstInvoices.List(sel, 4))
    
    ReadInvoiceTotals invNum, total, paid, credit
    due = Round(total - credit, 2)
    If due < 0 Then due = 0
    remaining = Round(due - paid, 2)
    If remaining < 0 Then remaining = 0
    
    amtText = Trim(InputBox( _
        "Invoice : " & invNum & vbNewLine & _
        "Customer: " & custName & vbNewLine & _
        "Balance : " & balanceText & vbNewLine & _
        "Date    : " & Format(paymentDate, "YYYY/MM/DD") & vbNewLine & vbNewLine & _
        "Enter the partial payment amount:", _
        "Apply Partial Payment", _
        Format(remaining, "0.00")))
    
    If amtText = "" Then Exit Sub
    
    If Not IsNumeric(amtText) Then
        MsgBox "'" & amtText & "' is not a valid amount.", vbExclamation, "Invalid Amount"
        Exit Sub
    End If
    
    payAmount = r2(amtText)
    If payAmount <= 0 Then
        MsgBox "Payment amount must be greater than zero.", vbExclamation, "Invalid Amount"
        Exit Sub
    End If
    
    ' Explicit overpayment confirmation (audit log will still capture the excess)
    If payAmount > remaining And remaining > 0 Then
        If MsgBox("Payment of " & FmtR(payAmount) & " EXCEEDS the remaining balance of " & FmtR(remaining) & "." & vbNewLine & vbNewLine & _
                  "The recorded amount will be capped at " & FmtR(remaining) & "," & vbNewLine & _
                  "and the excess (" & FmtR(payAmount - remaining) & ") will be noted in the audit log." & vbNewLine & vbNewLine & _
                  "Continue?", _
                  vbExclamation + vbYesNo, "Overpayment") <> vbYes Then Exit Sub
    End If
    
    ApplyPartialPayment invNum, payAmount, paymentDate
    
    LoadUnpaidInvoices
    UpdateButtonStates
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

