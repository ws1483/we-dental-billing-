Attribute VB_Name = "modFormulas"
Option Explicit

' =====================================================
' ===== SHARED FOOTER LAYOUT CONSTANTS =====
' =====================================================
Public Const QI_FOOTER_START_ROW As Long = 41
Public Const QI_FOOTER_END_ROW   As Long = 48

Public Const SHARED_FOOTER_ROW_COUNT As Long = 9

Private Const FOOTER_ACCENT_R As Long = 46
Private Const FOOTER_ACCENT_G As Long = 135
Private Const FOOTER_ACCENT_B As Long = 176

' Money format that hides zero and empty-string values.
' Trailing semicolon = "show nothing for zero".
Private Const MONEY_FORMAT_HIDE_ZERO As String = """R""#,##0.00;-""R""#,##0.00;"

' =====================================================
' ===== FOOTER ROW DEFINITION (table-driven layout) =====
' =====================================================
Private Type FooterRow
    relativeRow As Long
    layout As Integer        ' 0 = full-width merge, 1 = split left/right, 2 = banner
    leftText As String
    rightText As String
    fontSize As Single
    isBold As Boolean
    isBanner As Boolean
    rowHeight As Single
End Type
' =====================================================
' ===== RESTORE HEADER LOOKUP FORMULAS =====
' =====================================================
' Customer-driven lookups in the header block.
'
' IMPORTANT: SafeSetFormula un-merges before writing. The header layout
' relies on horizontal merges (C:D for the customer text fields), so we
' re-merge each target range AFTER the formula is written.
'
' Layout (from working reference):
'   C7  : Customer code (user picks here, dropdown) - NOT merged
'   C8:D8 : Practice (lookup from Customers!B) - merged, centered
'   C9:D9 : Address (lookup from Customers!C) - merged, centered
'   D11 : Postcode (lookup from Customers!D) - merged C11:D11, centered
'   G7  : Date (manual)
'   G8  : Quote/Invoice number (manual)
'   G9  : Cust ID (lookup from Customers!J)
'   G10 : VAT No (lookup from Customers!K)
'   G11 : Email (lookup, was previously cleared - now restored)

Public Sub RestoreHeaderLookups(ws As Worksheet)
    On Error Resume Next
    
    ' --- Cust ID (G9) ---
    SafeSetFormula ws.Range("G9"), _
        "=IF(C7="""","""",IFERROR(INDEX(Customers!$J:$J,MATCH(C7,Customers!$A:$A,0)),""""))"
    
    ' --- Practice name (C8) - merged C8:D8 ---
    SafeSetFormula ws.Range("C8"), _
        "=IF(C7="""","""",IFERROR(INDEX(Customers!$B:$B,MATCH(C7,Customers!$A:$A,0)),""""))"
    SafeMergeRange ws.Range("C8:D8"), True
    
    ' --- Address (C9) - merged C9:D9 ---
    SafeSetFormula ws.Range("C9"), _
        "=IF(C7="""","""",IFERROR(INDEX(Customers!$C:$C,MATCH(C7,Customers!$A:$A,0)),""""))"
    SafeMergeRange ws.Range("C9:D9"), True
    
    ' --- Postcode (C11) - merged C11:D11 ---
    ' Original code wrote the formula to D11; that worked because D11 was the
    ' left-most non-merged cell at the time. With the merge restored, the
    ' formula must live in the top-left cell of the merged range, i.e. C11.
    SafeSetFormula ws.Range("C11"), _
        "=IF(C7="""","""",IFERROR(INDEX(Customers!$D:$D,MATCH(C7,Customers!$A:$A,0)),""""))"
    SafeMergeRange ws.Range("C11:D11"), True
    
    ' --- VAT Number (G10) ---
    SafeSetFormula ws.Range("G10"), _
        "=IF(C7="""","""",IFERROR(INDEX(Customers!$K:$K,MATCH(C7,Customers!$A:$A,0)),""""))"
    
    ' --- Email (G11) - was previously cleared. Restore lookup. ---
    SafeSetFormula ws.Range("G11"), _
        "=IF(C7="""","""",IFERROR(INDEX(Customers!$E:$E,MATCH(C7,Customers!$A:$A,0)),""""))"
    
    On Error GoTo 0
End Sub

' Merges a range and applies the centering used by the header layout.
' Safe to call on a range that's already merged the same way.
Private Sub SafeMergeRange(rng As Range, ByVal centerHorizontally As Boolean)
    On Error Resume Next
    
    ' Un-merge first if there's a stale merge that's not the same shape we want
    If rng.MergeCells Then
        ' Only re-merge if the existing merged area differs from rng
        If rng.MergeArea.Address <> rng.Address Then
            rng.UnMerge
        End If
    End If
    
    If Not rng.MergeCells Then
        rng.Merge
    End If
    
    If centerHorizontally Then
        rng.HorizontalAlignment = xlCenter
    End If
    rng.VerticalAlignment = xlCenter
    
    On Error GoTo 0
End Sub

' =====================================================
' ===== SAFE FORMULA SETTER (locale-aware) =====
' =====================================================
' 1. Try .Formula (English locale, the canonical form)
' 2. Fall back to .FormulaLocal with comma decimals + semicolon separators
' 3. Verify the result was actually accepted as a formula

Public Sub SafeSetFormula(rng As Range, ByVal formulaText As String)
    On Error Resume Next
    
    If rng.MergeCells Then rng.UnMerge
    rng.ClearContents
    
    ' --- Attempt 1: English (.Formula) ---
    Err.Clear
    rng.Formula = formulaText
    If Err.Number = 0 Then
        If Left$(CStr(rng.Formula), 1) = "=" Then
            On Error GoTo 0
            Exit Sub
        End If
    End If
    
    ' --- Attempt 2: Locale (.FormulaLocal) with translated separators ---
    Err.Clear
    Dim localFormula As String
    localFormula = Replace(formulaText, ",", ";")
    localFormula = Replace(localFormula, ".", ",")
    rng.FormulaLocal = localFormula
    If Err.Number = 0 Then
        If Left$(CStr(rng.FormulaLocal), 1) = "=" Then
            On Error GoTo 0
            Exit Sub
        End If
    End If
    
    ' --- Attempt 3: surrender, store 0 so user notices something is off ---
    Err.Clear
    rng.Value = 0
    
    On Error GoTo 0
End Sub

' Returns the input with any comma decimal forced to a dot.
' Critical: without this, "0,15" from VatRateStr() on non-English
' locales becomes "0;15" after the comma->semicolon pass in
' SafeSetFormula's fallback path, mangling the formula.
Public Function NormaliseDecimal(ByVal s As String) As String
    NormaliseDecimal = Replace(s, ",", ".")
End Function

' Tiny helper that safely writes a plain-text label into a (possibly merged) cell.
Private Sub SafeWriteLabel(rng As Range, ByVal labelText As String)
    On Error Resume Next
    If rng.MergeCells Then rng.UnMerge
    rng.Value = labelText
    On Error GoTo 0
End Sub

' =====================================================
' ===== LINE-ITEM FORMULAS (Quote / Invoice rows 14-31) =====
' =====================================================
' Workflow:
'   User picks Description from dropdown in column D.
'   Code (B), Z-code (C), Excl price (E), VAT (F), Incl (G), Total (H)
'   all derive from that.
'   Every row returns "" (blank) when column D is empty so unused rows
'   stay visually clean.
'
' PriceList layout (verified):
'   A = Item Code
'   B = Description (the lookup key)
'   C = Z Code
'   F = Current Price Excl
'   G = Current Price Incl

Private Sub WriteLineItemFormulas(ws As Worksheet)
    Dim i As Integer
    Dim vat As String
    Dim cell As Range
    
    ' Force dot decimal regardless of locale so SafeSetFormula's
    ' comma->semicolon pass doesn't eat the VAT rate.
    vat = NormaliseDecimal(VatRateStr())   ' "0.15"
    
    For i = START_ROW To END_ROW
        
        ' --- B: Item Code (lookup) ---
        SafeSetFormula ws.Cells(i, 2), _
            "=IF($D" & i & "="""","""",IFERROR(INDEX(PriceList!$A:$A,MATCH($D" & i & ",PriceList!$B:$B,0)),""""))"
        
        ' --- C: Z Code (lookup) ---
        SafeSetFormula ws.Cells(i, 3), _
            "=IF($D" & i & "="""","""",IFERROR(INDEX(PriceList!$C:$C,MATCH($D" & i & ",PriceList!$B:$B,0)),""""))"
        
        ' --- E: Price Excl (lookup) ---
        SafeSetFormula ws.Cells(i, 5), _
            "=IF($D" & i & "="""","""",IFERROR(ROUND(INDEX(PriceList!$F:$F,MATCH($D" & i & ",PriceList!$B:$B,0)),2),""""))"
        
        ' --- F: VAT amount on line --- = Excl * VAT_RATE
        SafeSetFormula ws.Cells(i, 6), _
            "=IF($E" & i & "="""","""",IFERROR(ROUND($E" & i & "*" & vat & ",2),""""))"
        
        ' --- G: Price Incl --- = Excl + VAT
        SafeSetFormula ws.Cells(i, 7), _
            "=IF($E" & i & "="""","""",IFERROR(ROUND($E" & i & "+$F" & i & ",2),""""))"
        
        ' --- H: Line Total --- = qty * price incl
        SafeSetFormula ws.Cells(i, 8), _
            "=IF(OR($A" & i & "="""",$G" & i & "=""""),"""",IFERROR(ROUND($G" & i & "*$A" & i & ",2),""""))"
        
        ' Format money cells: currency with zero/empty hidden
        For Each cell In ws.Range("E" & i & ",F" & i & ",G" & i & ",H" & i)
            cell.NumberFormat = MONEY_FORMAT_HIDE_ZERO
        Next cell
    Next i
End Sub

' =====================================================
' ===== SUMMARY FORMULAS (Quote / Invoice rows 32-36) =====
' =====================================================
' Layout (per modConfig):
'   C32 (ROW_DISCOUNT_LABEL)  = Discount %     (user input)
'   G33 / H33 = "Discount"     / discount amount
'   G34 / H34 = "Sub-Total"    / line totals - discount, divided by 1+VAT_RATE
'   G35 / H35 = "VAT"          / VAT amount
'   G36 / H36 = "Total Due"    / Sub-Total + VAT
'
' Every value formula returns "" (blank) when COUNTA(D14:D31)=0
' so an empty document shows no totals.

Private Sub WriteSummaryFormulas(ws As Worksheet)
    Dim vat As String, vatPlusOne As String
    Dim cell As Range
    Dim r As Long
    
    ' Force dot decimals so the locale-conversion fallback doesn't mangle them.
    vat = NormaliseDecimal(VatRateStr())           ' "0.15"
    vatPlusOne = NormaliseDecimal(OneVatRateStr()) ' "1.15"
    
    On Error Resume Next
    
    ' Discount % cell: user-editable percentage
    ws.Range(CELL_DISCOUNT_PCT).NumberFormat = "0%"
    
    ' --- Step 1: clear stray cells in footer block F32:H36 ---
    For r = 32 To 36
        ws.Range("F" & r).ClearContents
        ws.Range("G" & r).ClearContents
        ws.Range("H" & r).ClearContents
    Next r
    
    ' --- Step 2: write labels in column G (rows 33-36) ---
    SafeWriteLabel ws.Cells(ROW_DISCOUNT_AMOUNT, 7), "Discount"
    SafeWriteLabel ws.Cells(ROW_SUBTOTAL, 7), "Sub-Total"
    SafeWriteLabel ws.Cells(ROW_VAT, 7), "VAT"
    SafeWriteLabel ws.Cells(ROW_TOTAL, 7), "Total Due"
    
    ' --- Step 3: write value formulas in column H (rows 33-36) ---
    
    ' Discount value: SUM(line totals) * Discount %
    SafeSetFormula ws.Cells(ROW_DISCOUNT_AMOUNT, 8), _
        "=IF(COUNTA(D" & START_ROW & ":D" & END_ROW & ")=0,"""",ROUND(SUM(H" & START_ROW & ":H" & END_ROW & ")*" & CELL_DISCOUNT_PCT & ",2))"
    
    ' Sub-Total (excl VAT) = (SUM(line totals) - discount) / (1 + VAT_RATE)
    SafeSetFormula ws.Cells(ROW_SUBTOTAL, 8), _
        "=IF(COUNTA(D" & START_ROW & ":D" & END_ROW & ")=0,"""",ROUND((SUM(H" & START_ROW & ":H" & END_ROW & ")-H" & ROW_DISCOUNT_AMOUNT & ")/" & vatPlusOne & ",2))"
    
    ' VAT amount = (SUM(line totals) - discount) - sub-total
    SafeSetFormula ws.Cells(ROW_VAT, 8), _
        "=IF(COUNTA(D" & START_ROW & ":D" & END_ROW & ")=0,"""",ROUND((SUM(H" & START_ROW & ":H" & END_ROW & ")-H" & ROW_DISCOUNT_AMOUNT & ")-H" & ROW_SUBTOTAL & ",2))"
    
    ' Total Due = Sub-Total + VAT
    SafeSetFormula ws.Cells(ROW_TOTAL, 8), _
        "=IF(COUNTA(D" & START_ROW & ":D" & END_ROW & ")=0,"""",ROUND(H" & ROW_SUBTOTAL & "+H" & ROW_VAT & ",2))"
    
    ' --- Step 4: formatting ---
    
    ' Value cells: currency, bold, hide zero/empty
    For Each cell In ws.Range("H" & ROW_DISCOUNT_AMOUNT & ",H" & ROW_SUBTOTAL & _
                              ",H" & ROW_VAT & ",H" & ROW_TOTAL)
        cell.NumberFormat = MONEY_FORMAT_HIDE_ZERO
        cell.Font.bold = True
    Next cell
    
    ' Label cells: bold + right-aligned
    For Each cell In ws.Range("G" & ROW_DISCOUNT_AMOUNT & ",G" & ROW_SUBTOTAL & _
                              ",G" & ROW_VAT & ",G" & ROW_TOTAL)
        cell.Font.bold = True
        cell.HorizontalAlignment = xlRight
    Next cell
    
    On Error GoTo 0
End Sub

' =====================================================
' ===== RESTORE QUOTE FORMULAS =====
' =====================================================

Public Sub RestoreQuoteFormulas()
    Dim ws As Worksheet
    
    If Not SheetExists(SHEET_QUOTE) Then
        ShowError "Quote sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_QUOTE)
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Restoring quote formulas..."
    
    RestoreHeaderLookups ws
    WriteLineItemFormulas ws
    WriteSummaryFormulas ws
    WriteSharedFooter ws, QI_FOOTER_START_ROW, True, "H"
    
    RestoreExcelState
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error restoring quote formulas: " & Err.Description
End Sub

' =====================================================
' ===== RESTORE INVOICE FORMULAS =====
' =====================================================

Public Sub RestoreInvoiceFormulas()
    Dim ws As Worksheet
    
    If Not SheetExists(SHEET_INVOICE) Then
        ShowError "Invoice sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
    
    On Error GoTo ErrorHandler
    TogglePerformance True
    SafeStatusBar "Restoring invoice formulas..."
    
    RestoreHeaderLookups ws
    WriteLineItemFormulas ws
    WriteSummaryFormulas ws
    WriteSharedFooter ws, QI_FOOTER_START_ROW, True, "H"
    
    RestoreExcelState
    Exit Sub

ErrorHandler:
    RestoreExcelState
    ShowError "Error restoring invoice formulas: " & Err.Description
End Sub

' =====================================================
' ===== ADD DESCRIPTION AND CUSTOMER DROPDOWNS =====
' =====================================================

Public Sub AddDescriptionDropdowns()
    Dim wsQuote As Worksheet, wsInvoice As Worksheet, wsPriceList As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, custLastRow As Long
    Dim i As Integer
    Dim listRange As String, custListRange As String
    
    On Error GoTo ErrorHandler
    
    If Not SheetExists(SHEET_PRICELIST) Then
        ShowError "PriceList sheet not found."
        Exit Sub
    End If
    
    Set wsPriceList = ThisWorkbook.Sheets(SHEET_PRICELIST)
    lastRow = wsPriceList.Cells(wsPriceList.rows.count, "B").End(xlUp).row
    
    If lastRow < 2 Then
        ShowError "No descriptions found in PriceList Column B."
        Exit Sub
    End If
    
    listRange = "=PriceList!$B$2:$B$" & lastRow
    
    If SheetExists(SHEET_QUOTE) Then
        Set wsQuote = ThisWorkbook.Sheets(SHEET_QUOTE)
        For i = START_ROW To END_ROW
            With wsQuote.Cells(i, 4).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listRange
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Next i
    End If
    
    If SheetExists(SHEET_INVOICE) Then
        Set wsInvoice = ThisWorkbook.Sheets(SHEET_INVOICE)
        For i = START_ROW To END_ROW
            With wsInvoice.Cells(i, 4).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listRange
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Next i
    End If
    
    If SheetExists(SHEET_CUSTOMERS) Then
        Set wsCust = ThisWorkbook.Sheets(SHEET_CUSTOMERS)
        custLastRow = wsCust.Cells(wsCust.rows.count, "A").End(xlUp).row
        
        If custLastRow >= 2 Then
            custListRange = "=Customers!$A$2:$A$" & custLastRow
            
            If SheetExists(SHEET_QUOTE) Then
                With ThisWorkbook.Sheets(SHEET_QUOTE).Range(CELL_CUSTOMER).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=custListRange
                    .IgnoreBlank = True
                    .InCellDropdown = True
                End With
            End If
            
            If SheetExists(SHEET_INVOICE) Then
                With ThisWorkbook.Sheets(SHEET_INVOICE).Range(CELL_CUSTOMER).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=custListRange
                    .IgnoreBlank = True
                    .InCellDropdown = True
                End With
            End If
        End If
    End If
    
    ShowInfo "Dropdowns added:" & vbNewLine & _
             "- Description (Col D): Rows " & START_ROW & " to " & END_ROW & vbNewLine & _
             "- Customer Name (" & CELL_CUSTOMER & "): From Customers sheet" & vbNewLine & _
             "Found " & (lastRow - 1) & " products.", "Complete"
    Exit Sub

ErrorHandler:
    ShowError "Error adding dropdowns: " & Err.Description
End Sub

Public Sub RemoveDescriptionDropdowns()
    Dim ws As Worksheet
    Dim i As Integer
    
    On Error Resume Next
    
    If SheetExists(SHEET_QUOTE) Then
        Set ws = ThisWorkbook.Sheets(SHEET_QUOTE)
        For i = START_ROW To END_ROW
            ws.Cells(i, 4).Validation.Delete
        Next i
        ws.Range(CELL_CUSTOMER).Validation.Delete
    End If
    
    If SheetExists(SHEET_INVOICE) Then
        Set ws = ThisWorkbook.Sheets(SHEET_INVOICE)
        For i = START_ROW To END_ROW
            ws.Cells(i, 4).Validation.Delete
        Next i
        ws.Range(CELL_CUSTOMER).Validation.Delete
    End If
    
    On Error GoTo 0
    ShowInfo "Dropdowns removed.", "Complete"
End Sub

Public Sub SetupAllFormulasAndDropdowns()
    RestoreQuoteFormulas
    RestoreInvoiceFormulas
    AddDescriptionDropdowns
    
    ShowInfo "All formulas, dropdowns and footer are set up." & vbNewLine & vbNewLine & _
             "Line items: Rows " & START_ROW & " to " & END_ROW & vbNewLine & _
             "VAT rate:  " & Format(VAT_RATE, "0.##%") & vbNewLine & _
             "Summary:" & vbNewLine & _
             "  " & CELL_DISCOUNT_PCT & " = Discount %    (user input)" & vbNewLine & _
             "  " & CELL_SUMMARY_DISC & " = Discount value (= SUM lines * " & CELL_DISCOUNT_PCT & ")" & vbNewLine & _
             "  " & CELL_SUMMARY_SUBTOT & " = Sub-Total EXCL VAT" & vbNewLine & _
             "  " & CELL_SUMMARY_VAT & " = VAT amount" & vbNewLine & _
             "  " & CELL_SUMMARY_TOTAL & " = Total Due (Sub-Total + VAT)" & vbNewLine & _
             "Footer: Rows " & QI_FOOTER_START_ROW & "-" & QI_FOOTER_END_ROW & " (cols A-H)" & vbNewLine & vbNewLine & _
             "Empty rows and zero totals are hidden automatically.", _
             "Setup Complete"
End Sub

' =====================================================
' ===== SHARED FOOTER (Quote / Invoice / Statement) =====
' =====================================================
' Banking-details footer block at row 41. Completely separate
' from the totals footer at rows 33-36.

Public Sub WriteSharedFooter(ws As Worksheet, _
                             ByVal startRow As Long, _
                             Optional ByVal compact As Boolean = False, _
                             Optional ByVal rightCol As String = "G")
    Dim footerColor As Long
    Dim leftEnd As String, rightStart As String
    Dim rows() As FooterRow
    Dim totalRows As Long
    Dim i As Long
    
    footerColor = RGB(FOOTER_ACCENT_R, FOOTER_ACCENT_G, FOOTER_ACCENT_B)
    
    Select Case UCase$(rightCol)
        Case "G"
            leftEnd = "C"
            rightStart = "E"
        Case "H"
            leftEnd = "D"
            rightStart = "E"
        Case Else
            leftEnd = "C"
            rightStart = "E"
    End Select
    
    rows = BuildFooterRows(compact)
    totalRows = UBound(rows) - LBound(rows) + 1
    
    On Error Resume Next
    
    With ws.Range("A" & startRow & ":" & rightCol & startRow + totalRows)
        .UnMerge
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.bold = False
    End With
    
    With ws.Range("A" & startRow & ":" & rightCol & startRow).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = footerColor
        .Weight = xlMedium
    End With
    
    For i = LBound(rows) To UBound(rows)
        RenderFooterRow ws, rows(i), startRow, footerColor, leftEnd, rightStart, rightCol
    Next i
    
    With ws.Range("A" & startRow + 2 & ":" & rightCol & startRow + totalRows - 1)
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With
    
    On Error GoTo 0
End Sub

Private Function BuildFooterRows(ByVal compact As Boolean) As FooterRow()
    Dim r() As FooterRow
    
    If compact Then
        ReDim r(0 To 7)
        
        SetFooterRow r(0), 0, 0, "", "", 0, False, False, 0
        SetFooterRow r(1), 1, 0, "Banking Details   We-Dental Prints. : First National Bank   Current Account: 63148517551   Branch: 252045", _
                                                                                                            "", 10, True, False, 20
        SetFooterRow r(2), 2, 0, "We-Dental Prints (PTY)ltd 2024 / 507911 / 07   Adress: 552 Duniet str Elarduspark Pretoria Gauteng South Africa", _
                                                                                                            "", 9, False, False, 0
        SetFooterRow r(3), 3, 0, "Lab Reg Number: We-Dental Laboratory's: TL0157     VAT Reg No:", "", 9, False, False, 0
        SetFooterRow r(4), 4, 1, "Ernst Vd Berg (079) 152-5444", "ernst@we-dental.co.za", 9, False, False, 0
        SetFooterRow r(5), 5, 1, "Wentzel Greyling (071) 353-6660", "Wentzel@we-dental.co.za", 9, False, False, 0
        SetFooterRow r(6), 6, 0, "", "", 9, False, False, 0
        SetFooterRow r(7), 7, 2, "www.we-dental.co.za", "", 11, True, True, 22
    Else
        ReDim r(0 To 9)
        
        SetFooterRow r(0), 0, 0, "", "", 0, False, False, 0
        SetFooterRow r(1), 1, 0, "Banking Details   We-Dental Prints. : First National Bank   Current Account: 63148517551   Branch: 252045", _
                                                                                                            "", 10, True, False, 20
        SetFooterRow r(2), 2, 0, "We-Dental Prints (PTY)ltd 2024 / 507911 / 07   Adress: 552 Duniet str Elarduspark Pretoria Gauteng South Africa", _
                                                                                                            "", 9, False, False, 0
        SetFooterRow r(3), 3, 0, "Lab Reg Number: We-Dental Laboratory's: TL0157", "", 9, False, False, 0
        SetFooterRow r(4), 4, 0, "VAT Reg No:", "", 9, False, False, 0
        SetFooterRow r(5), 5, 0, "", "", 9, False, False, 0
        SetFooterRow r(6), 6, 1, "Ernst Vd Berg (079) 152-5444", "ernst@we-dental.co.za", 9, False, False, 0
        SetFooterRow r(7), 7, 1, "Wentzel Greyling (071) 353-6660", "Wentzel@we-dental.co.za", 9, False, False, 0
        SetFooterRow r(8), 8, 0, "", "", 9, False, False, 0
        SetFooterRow r(9), 9, 2, "www.we-dental.co.za", "", 11, True, True, 22
    End If
    
    BuildFooterRows = r
End Function

Private Sub SetFooterRow(ByRef target As FooterRow, _
                          ByVal relRow As Long, _
                          ByVal layout As Integer, _
                          ByVal leftTxt As String, _
                          ByVal rightTxt As String, _
                          ByVal fontSz As Single, _
                          ByVal bold As Boolean, _
                          ByVal banner As Boolean, _
                          ByVal rowH As Single)
    target.relativeRow = relRow
    target.layout = layout
    target.leftText = leftTxt
    target.rightText = rightTxt
    target.fontSize = fontSz
    target.isBold = bold
    target.isBanner = banner
    target.rowHeight = rowH
End Sub

Private Sub RenderFooterRow(ws As Worksheet, _
                             ByRef fr As FooterRow, _
                             ByVal startRow As Long, _
                             ByVal footerColor As Long, _
                             ByVal leftEnd As String, _
                             ByVal rightStart As String, _
                             ByVal rightCol As String)
    Dim r As Long
    r = startRow + fr.relativeRow
    
    If fr.leftText = "" And fr.rightText = "" And fr.layout = 0 And Not fr.isBanner Then
        If fr.rowHeight > 0 Then ws.rows(r).rowHeight = fr.rowHeight
        Exit Sub
    End If
    
    Select Case fr.layout
        Case 0
            ws.Range("A" & r & ":" & rightCol & r).Merge
            With ws.Cells(r, "A")
                .Value = fr.leftText
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.bold = fr.isBold
                .Font.Size = fr.fontSize
                .Font.Color = footerColor
            End With
            
            If fr.isBold And fr.fontSize >= 10 Then
                With ws.Range("A" & r & ":" & rightCol & r)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Color = footerColor
                    .Borders(xlEdgeTop).Weight = xlThin
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Color = footerColor
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).Color = footerColor
                    .Borders(xlEdgeLeft).Weight = xlThin
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).Color = footerColor
                    .Borders(xlEdgeRight).Weight = xlThin
                End With
            End If
            
        Case 1
            ws.Range("A" & r & ":" & leftEnd & r).Merge
            With ws.Cells(r, "A")
                .Value = fr.leftText
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Font.bold = fr.isBold
                .Font.Size = fr.fontSize
                .Font.Color = footerColor
            End With
            
            ws.Range(rightStart & r & ":" & rightCol & r).Merge
            With ws.Cells(r, rightStart)
                .Value = fr.rightText
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .Font.bold = fr.isBold
                .Font.Size = fr.fontSize
                .Font.Color = footerColor
            End With
            
        Case 2
            ws.Range("A" & r & ":" & rightCol & r).Merge
            With ws.Cells(r, "A")
                .Value = fr.leftText
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.bold = fr.isBold
                .Font.Size = fr.fontSize
                .Font.Color = RGB(255, 255, 255)
            End With
            ws.Range("A" & r & ":" & rightCol & r).Interior.Color = footerColor
    End Select
    
    If fr.rowHeight > 0 Then ws.rows(r).rowHeight = fr.rowHeight
End Sub

