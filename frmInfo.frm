VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time Zone Info"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   5295
         TabIndex        =   1
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   495
         Index           =   0
         Left            =   2880
         TabIndex        =   4
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   495
         Index           =   1
         Left            =   2880
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   2
         Top             =   3720
         Width           =   2775
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   2520
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Frame1.Caption = CurrentTZI.DisplayName
  Label1(0) = "Local time at computer Time Zone"
  Label1(1) = "Coordinated Universal Time (UTC)"
  Label1(2) = "Local time at choosen Time Zone"
  Picture1.AutoRedraw = True
  Timer1.Interval = 1000
  Timer1.Enabled = True
  FillText
  Timer1_Timer
End Sub

Private Function SignStr(ByVal sng As Single) As String
  Dim s As String
  s = CStr(sng)
  If Left$(s, 1) <> "-" Then s = "+" & s
  SignStr = s
End Function

Private Function TranslateDay(ByVal nDayOfWeek&, ByVal nDay&) As String
  Dim sReturn$
  sReturn = "The "
  Select Case nDay
    Case 1: sReturn = sReturn & "First "
    Case 2: sReturn = sReturn & "Second "
    Case 3: sReturn = sReturn & "Third "
    Case 4: sReturn = sReturn & "Fourth "
    Case 5: sReturn = sReturn & "Last "
  End Select
  Select Case nDayOfWeek
    Case 0: sReturn = sReturn & "Sunday"
    Case 1: sReturn = sReturn & "Monday"
    Case 2: sReturn = sReturn & "Tuesday"
    Case 3: sReturn = sReturn & "Wednesday"
    Case 4: sReturn = sReturn & "Thursday"
    Case 5: sReturn = sReturn & "Friday"
    Case 6: sReturn = sReturn & "Saturday"
  End Select
  TranslateDay = sReturn & " In"
End Function

Private Function GetMonth(ByVal nMonth&) As String
  Select Case nMonth
    Case 1: GetMonth = "January"
    Case 2: GetMonth = "February"
    Case 3: GetMonth = "March"
    Case 4: GetMonth = "April"
    Case 5: GetMonth = "May"
    Case 6: GetMonth = "June"
    Case 7: GetMonth = "July"
    Case 8: GetMonth = "August"
    Case 9: GetMonth = "September"
    Case 10: GetMonth = "October"
    Case 11: GetMonth = "November"
    Case 12: GetMonth = "December"
  End Select
End Function


Private Function SysDate() As Date
   Dim st As SYSTEMTIME
   Call GetSystemTime(st)
   SysDate = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

Private Sub FillText()
  Picture1.Cls
  With CurrentTZI
       Picture1.Print "Normal Bias ", , SignStr(.Bias / 60) & " hour(s) to convert local time to UTC"
       Picture1.Print
       Picture1.Print "Standard Name ", .StandardName
       Picture1.Print "Standard Bias ", SignStr(.StandardBias / 60) & " hour(s) to add to Normal Bias"
       Picture1.Print
       If .DaylightDate.wMonth = 0 Then
          Picture1.Print "No DayLight difference at this zone"
          Exit Sub
       End If
       Picture1.Print "Daylight Name  ", .DaylightName
       Picture1.Print "DayLight Bias ", SignStr(.DaylightBias / 60) & " hour(s) to add to Normal Bias"
  End With
  With CurrentTZI.DaylightDate
       If .wYear Then
          Picture1.Print "Daylight Begins On ", DateSerial(.wYear, .wMonth, .wDay) & " At "; TimeSerial(.wHour, .wMinute, .wSecond)
       Else
          Picture1.Print "Daylight Begins On ", TranslateDay(.wDayOfWeek, .wDay) & " " & GetMonth(.wMonth) & " At " & TimeSerial(.wHour, .wMinute, .wSecond)
          Picture1.Print , , "(This year - " & Format(WeekDayToDate(Year(Now), .wMonth, .wDayOfWeek, .wDay), "dd mmmm yyyy)")
       End If
  End With
  With CurrentTZI.StandardDate
       If .wYear Then
          Picture1.Print "Daylight Ends On ", DateSerial(.wYear, .wMonth, .wDay) & " At "; TimeSerial(.wHour, .wMinute, .wSecond)
       Else
          Picture1.Print "Daylight Ends On ", TranslateDay(.wDayOfWeek, .wDay) & " " & GetMonth(.wMonth) & " At " & TimeSerial(.wHour, .wMinute, .wSecond)
          Picture1.Print , , "(This year - " & Format(WeekDayToDate(Year(Now), .wMonth, .wDayOfWeek, .wDay), "dd mmmm yyyy)")
       End If
  End With
End Sub

Private Sub Timer1_Timer()
  Dim d As Date
  d = Now
  Label2(0) = Format(d, "Long date") & vbCrLf & Format(d, "hh:mm:ss")
  d = SysDate
  Label2(1) = Format(d, "Long date") & vbCrLf & Format(d, "hh:mm:ss")
  d = UTCToLocalDate(SysDate, CurrentTZI)
  Label2(2) = Format(d, "Long date") & vbCrLf & Format(d, "hh:mm:ss")
End Sub

