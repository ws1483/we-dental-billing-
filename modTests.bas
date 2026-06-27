Attribute VB_Name = "modTests"
Option Explicit

' =====================================================
' ===== UNIT TESTS =====
' =====================================================
' Run RunAllTests from Alt+F8 to verify the math/parsing
' helpers haven't regressed. Test results land in a popup
' AND in the Immediate window (Ctrl+G to view).
'
' Tests are deliberately lightweight - no external test
' framework, just Debug.Print + a pass/fail counter.
' =====================================================

Private mPassCount As Long
Private mFailCount As Long
Private mFailLog   As String

Public Sub RunAllTests()
    mPassCount = 0
    mFailCount = 0
    mFailLog = ""
    
    Debug.Print String(50, "=")
    Debug.Print "Running tests - " & Now
    Debug.Print String(50, "=")
    
    Test_r2_RoundsToTwoDp
    Test_r2_HandlesNonNumeric
    Test_FmtR_StandardFormat
    Test_ExtractSequence_ValidPrefixedInput
    Test_ExtractSequence_RejectsWrongPrefix
    Test_ExtractSequence_HandlesGarbage
    Test_VatRate_Math
    Test_AgingBuckets_Boundaries
    Test_IsUnpaidStatus_KnownStatuses
    Test_CleanFileName_StripsInvalidChars
    Test_NormaliseName_TrimAndCollapse
    
    Debug.Print String(50, "-")
    Debug.Print "Passed: " & mPassCount & "   Failed: " & mFailCount
    Debug.Print String(50, "=")
    
    Dim msg As String
    msg = "Tests complete" & vbNewLine & vbNewLine & _
          "Passed: " & mPassCount & vbNewLine & _
          "Failed: " & mFailCount
    
    If mFailCount > 0 Then
        msg = msg & vbNewLine & vbNewLine & "Failures:" & vbNewLine & mFailLog
        MsgBox msg, vbCritical, "Test Results"
    Else
        MsgBox msg & vbNewLine & vbNewLine & "All green " & ChrW(10004), vbInformation, "Test Results"
    End If
End Sub

' =====================================================
' ===== ASSERTION HELPERS =====
' =====================================================

Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If CStr(expected) = CStr(actual) Then
        mPassCount = mPassCount + 1
        Debug.Print "  PASS  " & testName
    Else
        mFailCount = mFailCount + 1
        Dim m As String
        m = testName & ": expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
        Debug.Print "  FAIL  " & m
        mFailLog = mFailLog & "  - " & m & vbNewLine
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal cond As Boolean)
    AssertEqual testName, True, cond
End Sub

Private Sub AssertClose(ByVal testName As String, ByVal expected As Double, _
                        ByVal actual As Double, Optional ByVal tolerance As Double = 0.005)
    If Abs(expected - actual) <= tolerance Then
        mPassCount = mPassCount + 1
        Debug.Print "  PASS  " & testName
    Else
        mFailCount = mFailCount + 1
        Dim m As String
        m = testName & ": expected ~" & expected & " got " & actual
        Debug.Print "  FAIL  " & m
        mFailLog = mFailLog & "  - " & m & vbNewLine
    End If
End Sub

' =====================================================
' ===== TESTS =====
' =====================================================

' --- r2 (money rounding to 2dp) ---
Private Sub Test_r2_RoundsToTwoDp()
    AssertEqual "r2 rounds 1.235 to 1.24 (banker's rounding)", 1.24, r2(1.235)
    AssertEqual "r2 rounds 1.234 down to 1.23", 1.23, r2(1.234)
    AssertEqual "r2 leaves 1.20 alone", 1.2, r2(1.2)
    AssertEqual "r2 of 0 = 0", 0, r2(0)
End Sub

Private Sub Test_r2_HandlesNonNumeric()
    AssertEqual "r2 of empty string = 0", 0, r2("")
    AssertEqual "r2 of Null = 0", 0, r2(Null)
    AssertEqual "r2 of 'abc' = 0", 0, r2("abc")
    AssertEqual "r2 of '12.5' = 12.5", 12.5, r2("12.5")
End Sub

' --- FmtR (display formatting) ---
Private Sub Test_FmtR_StandardFormat()
    AssertEqual "FmtR formats 1234.5 correctly", "R 1,234.50", FmtR(1234.5)
    AssertEqual "FmtR formats 0", "R 0.00", FmtR(0)
    AssertEqual "FmtR rounds 1.239", "R 1.24", FmtR(1.239)
End Sub

' --- ExtractSequence (parses INV-0042 -> 42, WA-7 -> 7) ---
Private Sub Test_ExtractSequence_ValidPrefixedInput()
    AssertEqual "ExtractSequence INV-0042 = 42", 42, ExtractSequenceProxy("INV-0042", "INV")
    AssertEqual "ExtractSequence WA-7 = 7", 7, ExtractSequenceProxy("WA-7", "WA")
    AssertEqual "ExtractSequence INV-100 = 100", 100, ExtractSequenceProxy("INV-100", "INV")
End Sub

Private Sub Test_ExtractSequence_RejectsWrongPrefix()
    AssertEqual "ExtractSequence rejects WA-1 when expecting INV", 0, ExtractSequenceProxy("WA-1", "INV")
    AssertEqual "ExtractSequence rejects QUOTE-1 when expecting WA", 0, ExtractSequenceProxy("QUOTE-1", "WA")
End Sub

Private Sub Test_ExtractSequence_HandlesGarbage()
    AssertEqual "ExtractSequence of empty = 0", 0, ExtractSequenceProxy("", "INV")
    AssertEqual "ExtractSequence of INV- (no num) = 0", 0, ExtractSequenceProxy("INV-", "INV")
    AssertEqual "ExtractSequence of INVABC = 0", 0, ExtractSequenceProxy("INVABC", "INV")
End Sub

' ExtractSequence is Private in modExternalTracker so we proxy it
' through what we can observe via GetNext... functions; alternatively
' temporarily change ExtractSequence to Public during a test run.
' For now we simulate the same logic locally so the test is self-contained.
Private Function ExtractSequenceProxy(ByVal textValue As String, ByVal expectedPrefix As String) As Long
    Dim s As String, dashPos As Long, foundPrefix As String
    
    On Error GoTo Failed
    
    s = Trim$(textValue)
    If s = "" Then Exit Function
    
    dashPos = InStr(1, s, "-")
    If dashPos = 0 Then Exit Function
    
    foundPrefix = Left$(s, dashPos - 1)
    
    If Len(expectedPrefix) > 0 Then
        If StrComp(foundPrefix, expectedPrefix, vbTextCompare) <> 0 Then Exit Function
    End If
    
    s = Mid$(s, dashPos + 1)
    If s = "" Or Not IsNumeric(s) Then Exit Function
    
    ExtractSequenceProxy = CLng(s)
    Exit Function

Failed:
    ExtractSequenceProxy = 0
End Function

' --- VAT math sanity ---
Private Sub Test_VatRate_Math()
    ' For a R 115.00 total incl VAT, subtotal should be 100, VAT 15
    Dim totalIncl As Double, sub_ As Double, vat As Double
    totalIncl = 115#
    sub_ = Round(totalIncl / (1 + VAT_RATE), 2)
    vat = Round(totalIncl - sub_, 2)
    
    AssertClose "VAT split: subtotal of R 115 = R 100", 100#, sub_
    AssertClose "VAT split: VAT of R 115 = R 15", 15#, vat
End Sub

' --- Aging bucket boundaries (mirrors AddToAgingBucket logic) ---
Private Sub Test_AgingBuckets_Boundaries()
    AssertEqual "Aging: day 0 = current", "current", BucketForAge(0)
    AssertEqual "Aging: day 30 = current", "current", BucketForAge(30)
    AssertEqual "Aging: day 31 = days30", "days30", BucketForAge(31)
    AssertEqual "Aging: day 60 = days30", "days30", BucketForAge(60)
    AssertEqual "Aging: day 61 = days60", "days60", BucketForAge(61)
    AssertEqual "Aging: day 90 = days60", "days60", BucketForAge(90)
    AssertEqual "Aging: day 91 = days90", "days90", BucketForAge(91)
    AssertEqual "Aging: day 1000 = days90", "days90", BucketForAge(1000)
End Sub

Private Function BucketForAge(ByVal d As Long) As String
    Select Case d
        Case 0 To 30:  BucketForAge = "current"
        Case 31 To 60: BucketForAge = "days30"
        Case 61 To 90: BucketForAge = "days60"
        Case Else:     BucketForAge = "days90"
    End Select
End Function

' --- IsUnpaidStatus (mirrors logic in modUnpaidInvoices) ---
Private Sub Test_IsUnpaidStatus_KnownStatuses()
    AssertTrue "Unpaid is unpaid", IsUnpaidStatusProxy("Unpaid")
    AssertTrue "Partial is unpaid", IsUnpaidStatusProxy("Partial")
    AssertTrue "Overdue is unpaid", IsUnpaidStatusProxy("Overdue")
    AssertEqual "Paid is NOT unpaid", False, IsUnpaidStatusProxy("Paid")
    AssertEqual "Cancelled is NOT unpaid", False, IsUnpaidStatusProxy("Cancelled")
    AssertEqual "Empty is NOT unpaid", False, IsUnpaidStatusProxy("")
    AssertEqual "Random is unpaid (default)", True, IsUnpaidStatusProxy("Pending Review")
End Sub

Private Function IsUnpaidStatusProxy(ByVal status As String) As Boolean
    Dim s As String
    s = LCase$(Trim(status))
    Select Case s
        Case ""
            IsUnpaidStatusProxy = False
        Case "paid", "cancelled", "canceled", "void", "voided", "credit", "credited", "refunded"
            IsUnpaidStatusProxy = False
        Case Else
            IsUnpaidStatusProxy = True
    End Select
End Function

' --- CleanFileName strips invalid chars ---
Private Sub Test_CleanFileName_StripsInvalidChars()
    AssertEqual "CleanFileName strips slashes", "DrSmith", CleanFileName("Dr/Smith")
    AssertEqual "CleanFileName strips colons", "DrSmith", CleanFileName("Dr:Smith")
    AssertEqual "CleanFileName strips asterisks", "DrSmith", CleanFileName("Dr*Smith")
    AssertEqual "CleanFileName collapses spaces", "Dr Smith", CleanFileName("Dr  Smith")
    AssertEqual "CleanFileName trims", "Dr Smith", CleanFileName("  Dr Smith  ")
End Sub

' --- NormaliseName (used by FindCustomerRow) ---
Private Sub Test_NormaliseName_TrimAndCollapse()
    ' We replicate the logic locally since NormaliseName is Private to modStatements.
    AssertEqual "Normalise replaces nbsp", "Dr Smith", NormaliseProxy("Dr" & Chr(160) & "Smith")
    AssertEqual "Normalise removes dots", "Dr Smith", NormaliseProxy("Dr. Smith")
    AssertEqual "Normalise collapses spaces", "Dr Smith", NormaliseProxy("Dr   Smith")
End Sub

Private Function NormaliseProxy(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Replace(t, Chr(160), " ")
    t = Replace(t, ".", "")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormaliseProxy = Trim(t)
End Function

