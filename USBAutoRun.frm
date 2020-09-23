VERSION 5.00
Begin VB.Form USBAutoRunFrm 
   Caption         =   "USBAutorun"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   Icon            =   "USBAutoRun.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   2685
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "USBAutoRunFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Drive1(26) As Boolean
Dim KnownDrive(26) As Boolean
Private Sub Form_Load()
 If App.PrevInstance Then End
 Drive1DotRefresh
 For i = 1 To 26
  KnownDrive(i) = Drive1(i)
 Next i
 Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
 On Error Resume Next
 Drive1DotRefresh
 For i = 4 To 26
  If Drive1(i) Then ' A USB Drive is currently inserted...
   If KnownDrive(i) = False Then ' Ahh - it just happened!
    KnownDrive(i) = True ' OK, so remember it's there.
    DriveLetter$ = Chr$(64 + i) & ":"
    AutoPlayFile$ = DriveLetter$ & "\AUTORUN.INF"
    If Dir$(AutoPlayFile$) <> "" Then
     Open AutoPlayFile$ For Input As #1
     While Not EOF(1)
      Line Input #1, A$
      If InStr(UCase$(A$), "OPEN=") Or InStr(UCase$(A$), "OPEN =") Then
        S = InStr(A$, "=")
        Program$ = Trim$(Right$(A$, Len(A$) - S))
        Program$ = DriveLetter$ & IIf(Left$(Program$, 1) <> "\", "\", "") & Program$
        Shell Program$, vbNormalFocus
        Close
        GoTo UpdateKnownDrive:
      End If
     Wend
     Close
    End If
    Exit Sub
   End If
   KnownDrive(i) = True
  End If
 Next i
UpdateKnownDrive:
 For i = 4 To 26
  KnownDrive(i) = Drive1(i)
 Next i
End Sub
Sub Drive1DotRefresh()
 On Error GoTo NoDrive:
 For i = 4 To 26
  DriveLetter$ = Chr$(64 + i)
  Drive1(i) = Dir$(DriveLetter$ & ":\AutoRun.inf") <> ""
NextDrive:
 Next i
 T = Timer + 0.1
 While T > Timer
  DoEvents
 Wend
 Exit Sub
NoDrive:
 Resume NextDrive:
End Sub
