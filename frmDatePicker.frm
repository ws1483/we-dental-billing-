VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "Select Date"
   ClientHeight    =   5232
   ClientLeft      =   12
   ClientTop       =   660
   ClientWidth     =   4980
   OleObjectBlob   =   "frmDatePicker.frx":0000
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================
' ===== DATE PICKER — MONTHLY CALENDAR GRID =====
' ===== (button-reuse optimization) =====
' =====================================================

Public SelectedDate As Date
Public Cancelled As Boolean

Private mYear As Integer
Private mMonth As Integer

' Persistent containers — built ONCE in UserForm_Initialize
Private mDayButtons() As MSForms.CommandButton
Private mDayHandlers() As clsDateButton
Private mControlsBuilt As Boolean

Private Const GRID_ROWS As Long = 6
Private Const GRID_COLS As Long = 7
Private Const TOTAL_CELLS As Long = GRID_ROWS * GRID_COLS  ' 42

Private Const DAY_BTN_W As Long = 30
Private Const DAY_BTN_H As Long = 22
Private Const LEFT_MARGIN As Long = 10
Private Const TOP_START As Long = 40

Private Const COLOR_TODAY As Long = 13428223   ' RGB(200, 230, 255) light blue
Private Const COLOR_NORMAL As Long = -2147483633 ' xl button face default

' =====================================================
' ===== LIFECYCLE =====
' =====================================================

Private Sub UserForm_Initialize()
    Me.caption = "Select Date"
    Me.Width = 260
    Me.Height = 290
    
    Cancelled = True
    SelectedDate = Date
    mYear = Year(Date)
    mMonth = Month(Date)
    
    If Not mControlsBuilt Then
        BuildDayHeaders
        BuildDayButtonsOnce
        mControlsBuilt = True
    End If
    
    RefreshCalendar
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Me.Hide
    End If
End Sub

' =====================================================
' ===== ONE-TIME: DAY-OF-WEEK HEADERS =====
' =====================================================

Private Sub BuildDayHeaders()
    Dim headers As Variant
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim existing As MSForms.Control
    
    headers = Array("Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")
    
    For i = 0 To 6
        ' If a previous instance left a header behind, reuse it
        Set existing = Nothing
        On Error Resume Next
        Set existing = Me.Controls("lblDow" & i)
        On Error GoTo 0
        
        If existing Is Nothing Then
            Set lbl = Me.Controls.Add("Forms.Label.1", "lblDow" & i)
        Else
            Set lbl = existing
        End If
        
        With lbl
            .caption = headers(i)
            .Top = TOP_START
            .Left = LEFT_MARGIN + (i * DAY_BTN_W)
            .Width = DAY_BTN_W
            .Height = 16
            .Font.bold = True
            .Font.Size = 8
            .TextAlign = fmTextAlignCenter
        End With
    Next i
End Sub

' =====================================================
' ===== ONE-TIME: BUILD ALL 42 DAY BUTTONS =====
' =====================================================
' Created once and reused forever. RefreshCalendar just
' flips visibility, captions, tags and colours.

Private Sub BuildDayButtonsOnce()
    Dim row As Integer, col As Integer
    Dim cellIndex As Integer
    Dim btnName As String
    Dim btn As MSForms.CommandButton
    Dim existing As MSForms.Control
    Dim handler As clsDateButton
    Dim topOffset As Long
    
    topOffset = TOP_START + 20
    
    ReDim mDayButtons(0 To TOTAL_CELLS - 1)
    ReDim mDayHandlers(0 To TOTAL_CELLS - 1)
    
    For row = 0 To GRID_ROWS - 1
        For col = 0 To GRID_COLS - 1
            cellIndex = row * GRID_COLS + col
            btnName = "btnDay" & Format(cellIndex, "00")
            
            ' Reuse if already on the form (e.g. from designer leftover)
            Set existing = Nothing
            On Error Resume Next
            Set existing = Me.Controls(btnName)
            On Error GoTo 0
            
            If existing Is Nothing Then
                Set btn = Me.Controls.Add("Forms.CommandButton.1", btnName)
            Else
                Set btn = existing
            End If
            
            With btn
                .caption = ""
                .Tag = ""
                .Top = topOffset + (row * DAY_BTN_H)
                .Left = LEFT_MARGIN + (col * DAY_BTN_W)
                .Width = DAY_BTN_W
                .Height = DAY_BTN_H
                .Font.Size = 9
                .Visible = False    ' will be flipped on each refresh
            End With
            
            Set mDayButtons(cellIndex) = btn
            
            Set handler = New clsDateButton
            handler.Init Me, btn
            Set mDayHandlers(cellIndex) = handler
        Next col
    Next row
End Sub

' =====================================================
' ===== REFRESH CALENDAR (cheap — no controls created) =====
' =====================================================
' Called on Prev / Next / SetInitialDate / Initialize.

Private Sub RefreshCalendar()
    Dim firstDay As Date
    Dim startDow As Integer
    Dim daysInMonth As Integer
    Dim row As Integer, col As Integer
    Dim cellIndex As Integer
    Dim dayNum As Integer
    Dim btn As MSForms.CommandButton
    Dim isToday As Boolean
    Dim lastVisibleCell As Integer
    Dim topOffset As Long
    
    On Error Resume Next
    Me.lblMonthYear.caption = MonthName(mMonth) & " " & mYear
    Me.btnToday.caption = "Today (" & Format(Date, "YYYY/MM/DD") & ")"
    On Error GoTo 0
    
    firstDay = DateSerial(mYear, mMonth, 1)
    startDow = Weekday(firstDay, vbSunday) - 1
    daysInMonth = Day(DateSerial(mYear, mMonth + 1, 0))
    
    dayNum = 1
    lastVisibleCell = -1
    
    For row = 0 To GRID_ROWS - 1
        For col = 0 To GRID_COLS - 1
            cellIndex = row * GRID_COLS + col
            Set btn = mDayButtons(cellIndex)
            
            If cellIndex >= startDow And dayNum <= daysInMonth Then
                isToday = (dayNum = Day(Date) And mMonth = Month(Date) And mYear = Year(Date))
                
                With btn
                    .caption = CStr(dayNum)
                    .Tag = CStr(dayNum)
                    .Visible = True
                    .Enabled = True
                    
                    If isToday Then
                        .BackColor = COLOR_TODAY
                        .Font.bold = True
                    Else
                        .BackColor = COLOR_NORMAL
                        .Font.bold = False
                    End If
                End With
                
                lastVisibleCell = cellIndex
                dayNum = dayNum + 1
            Else
                With btn
                    .caption = ""
                    .Tag = ""
                    .Visible = False
                End With
            End If
        Next col
    Next row
    
    ' Position the Today button just below the last used row of the grid
    Dim lastRowUsed As Long
    If lastVisibleCell >= 0 Then
        lastRowUsed = lastVisibleCell \ GRID_COLS
    Else
        lastRowUsed = 0
    End If
    
    topOffset = TOP_START + 20
    
    Dim todayTop As Long
    todayTop = topOffset + ((lastRowUsed + 1) * DAY_BTN_H) + 8
    
    On Error Resume Next
    Me.btnToday.Top = todayTop
    Me.Height = todayTop + 60
    On Error GoTo 0
End Sub

' =====================================================
' ===== NAVIGATION (designer button events) =====
' =====================================================

Private Sub btnPrev_Click()
    mMonth = mMonth - 1
    If mMonth < 1 Then
        mMonth = 12
        mYear = mYear - 1
    End If
    RefreshCalendar
End Sub

Private Sub btnNext_Click()
    mMonth = mMonth + 1
    If mMonth > 12 Then
        mMonth = 1
        mYear = mYear + 1
    End If
    RefreshCalendar
End Sub

Private Sub btnToday_Click()
    SelectedDate = Date
    Cancelled = False
    Me.Hide
End Sub

' =====================================================
' ===== DAY CLICK CALLBACK (called by clsDateButton) =====
' =====================================================

Public Sub DayClicked(ByVal dayNum As Integer)
    SelectedDate = DateSerial(mYear, mMonth, dayNum)
    Cancelled = False
    Me.Hide
End Sub

' =====================================================
' ===== PUBLIC: SET INITIAL MONTH/YEAR =====
' =====================================================

Public Sub SetInitialDate(ByVal d As Date)
    If d > 0 And IsDate(d) Then
        mYear = Year(d)
        mMonth = Month(d)
        SelectedDate = d
        
        ' If form is already open, refresh immediately
        If mControlsBuilt Then RefreshCalendar
    End If
End Sub
