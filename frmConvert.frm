VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Index           =   1
         Left            =   4440
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   1575
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Index           =   0
         Left            =   4440
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TZI1(1) As LOCALE_TIME_ZONE_INFORMATION
Dim nScroll1 As Integer, nScroll2 As Integer

Private Sub Combo1_Click(Index As Integer)
  TZI1(Index) = LocTZI(Combo1(Index).ListIndex)
  ShowNewTime
End Sub

Private Sub Command1_Click(Index As Integer)
  CurrentTZI = TZI1(Index)
  frmInfo.Show vbModal
End Sub

Private Sub Form_Load()
  Dim sRegKeyTZI As String, sCurZone As String, n As Integer, i As Integer
  If IsNT Then
     sRegKeyTZI = "Software\Microsoft\Windows NT\CurrentVersion\Time Zones"
  Else
     sRegKeyTZI = "Software\Microsoft\Windows\CurrentVersion\Time Zones"
  End If
  If Not GetTZICollection(sRegKeyTZI) Then
     MsgBox "Unable to locate Time Zones information in Registry under the key: " & vbCrLf & sRegKeyTZI
     Unload Me
     End
  End If
  sCurZone = GetRegValueStr("System\CurrentControlSet\Control\TimeZoneInformation", "StandardName")
  Caption = "Time Zones Conversion"
  Frame1.Caption = "From zone"
  Frame2.Caption = "To zone"
  Command1(0).Caption = "&Info"
  Command1(1).Caption = "I&nfo"
  For i = 0 To UBound(LocTZI)
      Combo1(0).AddItem LocTZI(i).DisplayName
      Combo1(1).AddItem LocTZI(i).DisplayName
      If LocTZI(i).StandardName = sCurZone Then n = i
  Next
  Text1 = Format(Now, "mmmm dd, yyyy")
  Text2 = Format(Now, "hh:mm:ss")
  Text1.Locked = True
  Text2.Locked = True
  nScroll1 = 4
  nScroll2 = 4
  VScroll1.Min = 0
  VScroll1.Max = 32000
  VScroll1.Value = 16000
  VScroll2.Min = 0
  VScroll2.Max = 32000
  VScroll2.Value = 16000
  Combo1(0).ListIndex = n
  Combo1(1).ListIndex = CInt(GetSetting("TimeZone", "Config", "ToZone", "0"))
  nScroll1 = 1
  nScroll2 = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call SaveSetting("TimeZone", "Config", "ToZone", CStr(Combo1(1).ListIndex))
End Sub

Private Sub Text1_Change()
   ShowNewTime
End Sub

Private Sub Text2_Change()
   ShowNewTime
End Sub

Private Sub Text1_GotFocus()
   Dim n1 As Long, n2 As Long
   If nScroll1 > 2 Then Exit Sub
   n1 = InStr(1, Text1, " ")
   n2 = InStr(1, Text1, ",")
   If nScroll1 = 0 Then
      Text1.SelStart = 0: Text1.SelLength = n1 - 1
   ElseIf nScroll1 = 2 Then
      Text1.SelStart = Len(Text1) - 4: Text1.SelLength = 4
   Else
      Text1.SelStart = Len(Text1) - 8: Text1.SelLength = 2
   End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim n1 As Long, n2 As Long, nCur As Long
   nCur = Text1.SelStart
   n1 = InStr(1, Text1, " ")
   n2 = InStr(1, Text1, ",")
   If nCur < n1 Then
      nScroll1 = 0
   ElseIf nCur > n2 Then
      nScroll1 = 2
   Else
      nScroll1 = 1
   End If
   Text1_GotFocus
End Sub

Private Sub Text2_GotFocus()
   If nScroll2 > 2 Then Exit Sub
   Text2.SelStart = nScroll2 * 3: Text2.SelLength = 2
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim nCur As Long
   nCur = Text2.SelStart
   nScroll2 = Int(nCur / 3)
   Text2_GotFocus
End Sub

Private Sub VScroll1_Change()
   Dim d As Date, k As Integer
   Static PrevValue As Integer
   If nScroll1 > 2 Then
      PrevValue = VScroll1.Value
      Exit Sub
   End If
   k = Sgn(VScroll1.Value - PrevValue)
   PrevValue = VScroll1.Value
   d = CDate(Text1)
   Select Case nScroll1
          Case 0
               Text1 = Format(DateSerial(Year(d), Month(d) - k, Day(d)), "mmmm dd, yyyy")
          Case 1
               Text1 = Format(DateSerial(Year(d), Month(d), Day(d) - k), "mmmm dd, yyyy")
          Case 2
               Text1 = Format(DateSerial(Year(d) - k, Month(d), Day(d)), "mmmm dd, yyyy")
   End Select
   Text1.SetFocus
End Sub

Private Sub VScroll2_Change()
   Dim d As Date, k As Integer
   Static PrevValue As Integer
   If nScroll2 > 2 Then
      PrevValue = VScroll2.Value
      Exit Sub
   End If
   k = Sgn(VScroll2.Value - PrevValue)
   PrevValue = VScroll2.Value
   d = CDate(Text2)
   Select Case nScroll2
          Case 0
               Text2 = Format(TimeSerial(Hour(d) - k, Minute(d), Second(d)), "hh:mm:ss")
          Case 1
               Text2 = Format(TimeSerial(Hour(d), Minute(d) - k, Second(d)), "hh:mm:ss")
          Case 2
               Text2 = Format(TimeSerial(Hour(d), Minute(d), Second(d) - k), "hh:mm:ss")
   End Select
   Text2.SetFocus
End Sub

Private Sub ShowNewTime()
  Dim d As Date
  On Error Resume Next
  d = CDate(Text1) + CDate(Text2)
  If Err Then Exit Sub
  d = LocalDateToUTC(d, TZI1(0))
  d = UTCToLocalDate(d, TZI1(1))
  Label1 = Format(d, "mmmm dd, yyyy    hh:mm:ss")
End Sub
