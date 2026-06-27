Attribute VB_Name = "modHelpers"
Option Explicit

' =====================================================
' ===== WINDOWS API: SHELL EXECUTE (open with default app) =====
' =====================================================
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

Private Const SW_SHOWNORMAL As Long = 1

' =====================================================
' ===== SHEET AND FILE HELPERS =====
' =====================================================

Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Function SheetExistsInWorkbook(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExistsInWorkbook = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Public Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

' Ensures the Settings sheet always exists with the expected layout.
' Safe to call repeatedly. Hidden as VeryHidden so users can't unhide via GUI.
Public Sub EnsureSettingsSheet()
    Dim ws As Worksheet
    
    If SheetExists(SHEET_SETTINGS) Then Exit Sub
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Name = SHEET_SETTINGS
    ws.Range("A1").Value = "QuoteCounter"
    ws.Range("A2").Value = "InvoiceCounter"
    ws.Range("A3").Value = "QuoteTrackingPath"
    ws.Range("A4").Value = "InvoiceTrackingPath"
    ws.Range("B1").Value = 0
    ws.Range("B2").Value = 0
    ws.Range("A1:A4").Font.bold = True
    ws.Columns("A:B").AutoFit
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' =====================================================
' ===== STRING HELPERS =====
' =====================================================

Public Function CleanFileName(ByVal strName As String) As String
    Dim invalidChars As Variant
    Dim i As Integer
    
    If Len(Trim(strName)) = 0 Then
        CleanFileName = ""
        Exit Function
    End If
    
    invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    
    For i = LBound(invalidChars) To UBound(invalidChars)
        strName = Replace(strName, invalidChars(i), "")
    Next i
    
    Do While InStr(strName, "  ") > 0
        strName = Replace(strName, "  ", " ")
    Loop
    
    CleanFileName = Trim(strName)
End Function

Public Function SafeString(val As Variant) As String
    On Error Resume Next
    If IsNull(val) Or IsEmpty(val) Then
        SafeString = ""
    Else
        SafeString = CStr(val)
    End If
    On Error GoTo 0
End Function

Public Function SafeDate(val As Variant) As Date
    On Error Resume Next
    If IsDate(val) Then
        SafeDate = CDate(val)
    Else
        SafeDate = Date
    End If
    On Error GoTo 0
End Function

Public Function SafeNumber(val As Variant) As Double
    On Error Resume Next
    If IsNumeric(val) Then
        SafeNumber = CDbl(val)
    Else
        SafeNumber = 0
    End If
    On Error GoTo 0
End Function

' =====================================================
' ===== MONEY ROUNDING / FORMATTING (used everywhere) =====
' =====================================================

Public Function r2(ByVal v As Variant) As Double
    r2 = Round(SafeNumber(v), 2)
End Function

' Standard "R 1,234.56" string used in popups and messages.
Public Function FmtR(ByVal v As Variant) As String
    FmtR = "R " & Format(r2(v), "#,##0.00")
End Function

' =====================================================
' ===== VAT RATE HELPERS (for formula strings) =====
' =====================================================
' Centralises the VAT rate so formulas don't hard-code 0.15 / 1.15.

Public Function VatRateStr() As String
    VatRateStr = Trim(CStr(VAT_RATE))
End Function

Public Function OneVatRateStr() As String
    OneVatRateStr = Trim(CStr(1 + VAT_RATE))
End Function

Public Function GetFileNameFromPath(filePath As String) As String
    Dim pos As Integer
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFileNameFromPath = Mid(filePath, pos + 1)
    Else
        GetFileNameFromPath = filePath
    End If
End Function

Public Function GetFolderFromPath(filePath As String) As String
    Dim pos As Integer
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFolderFromPath = Left(filePath, pos)
    Else
        GetFolderFromPath = ""
    End If
End Function

' =====================================================
' ===== PERFORMANCE / STATE HELPERS =====
' =====================================================

Public Sub TogglePerformance(turnOn As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = Not turnOn
    Application.EnableEvents = Not turnOn
    Application.Calculation = IIf(turnOn, xlCalculationManual, xlCalculationAutomatic)
    If Not turnOn Then SafeStatusBar False
    On Error GoTo 0
End Sub

Public Sub RestoreExcelState()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    SafeStatusBar False
    On Error GoTo 0
End Sub

Public Sub SafeStatusBar(ByVal msg As Variant)
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

' =====================================================
' ===== FILE OPEN (uses ShellExecute, future-proof) =====
' =====================================================
' Opens any file with its registered default app.
' Returns True on success, shows a friendly MsgBox on failure.

Public Function OpenWithDefaultApp(ByVal filePath As String) As Boolean
    #If VBA7 Then
        Dim ret As LongPtr
    #Else
        Dim ret As Long
    #End If
    
    OpenWithDefaultApp = False
    
    If Trim(filePath) = "" Then Exit Function
    If Not FileExists(filePath) Then
        ShowError "File not found:" & vbNewLine & vbNewLine & filePath
        Exit Function
    End If
    
    On Error Resume Next
    ret = ShellExecute(0&, "open", filePath, vbNullString, vbNullString, SW_SHOWNORMAL)
    On Error GoTo 0
    
    ' ShellExecute returns > 32 on success.
    If ret > 32 Then
        OpenWithDefaultApp = True
    Else
        ShowError "Could not open file (Windows error code " & ret & "):" & vbNewLine & vbNewLine & filePath
    End If
End Function

' =====================================================
' ===== DIALOG HELPERS =====
' =====================================================

Public Function GetSavePath(defaultFileName As String, fileFilter As String, dialogTitle As String) As String
    Dim savePath As Variant
    
    On Error Resume Next
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        fileFilter:=fileFilter, _
        title:=dialogTitle)
    On Error GoTo 0
    
    If VarType(savePath) = vbBoolean Then
        GetSavePath = ""
    Else
        GetSavePath = CStr(savePath)
    End If
End Function

Public Function GetOpenPath(fileFilter As String, dialogTitle As String) As String
    Dim openPath As Variant
    
    On Error Resume Next
    openPath = Application.GetOpenFilename( _
        fileFilter:=fileFilter, _
        title:=dialogTitle)
    On Error GoTo 0
    
    If VarType(openPath) = vbBoolean Then
        GetOpenPath = ""
    Else
        GetOpenPath = CStr(openPath)
    End If
End Function

Public Function GetFolderPath(dialogTitle As String) As String
    Dim folderPath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = dialogTitle
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
            If Right(folderPath, 1) <> "\" Then
                folderPath = folderPath & "\"
            End If
            GetFolderPath = folderPath
        Else
            GetFolderPath = ""
        End If
    End With
End Function

' =====================================================
' ===== MESSAGE HELPERS =====
' =====================================================

Public Sub ShowError(message As String, Optional title As String = "Error")
    MsgBox message, vbCritical, title
End Sub

Public Sub ShowInfo(message As String, Optional title As String = "Information")
    MsgBox message, vbInformation, title
End Sub

Public Sub ShowWarning(message As String, Optional title As String = "Warning")
    MsgBox message, vbExclamation, title
End Sub

Public Function AskYesNo(message As String, Optional title As String = "Confirm") As Boolean
    AskYesNo = (MsgBox(message, vbQuestion + vbYesNo, title) = vbYes)
End Function

