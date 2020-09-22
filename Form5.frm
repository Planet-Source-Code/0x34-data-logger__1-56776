VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "0x34's Device PING Tester"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Timer Timer4 
      Interval        =   150
      Left            =   3600
      Top             =   -120
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      InputLen        =   1
      RThreshold      =   1
      EOFEnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.CommandButton RunTest 
      BackColor       =   &H8000000B&
      Caption         =   "Run Test"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   4080
      Top             =   -120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "RS232 Device sends per second ="
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CNT As Long
Dim BVD As Integer

Private Sub Done_Click()
If MSComm2.PortOpen = True Then
    MSComm2.PortOpen = False
End If
Unload Me
End Sub

Private Sub Form_Load()
Timer4.Enabled = False
Label2 = ""
Timer3.Enabled = False
CNT = 0
BVD = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Timer4.Enabled = True Then Cancel = 1
End Sub

Private Sub MSComm2_OnComm()
CNT = CNT + 1
End Sub

Private Sub RunTest_Click()
    CNT = 0
    BVD = 0
    Call Comm
    RunTest.BackColor = vbRed
    RunTest.Caption = "TESTING"
    Timer3.Enabled = True
    Timer4.Enabled = True
    Label2 = ""
    Done.Enabled = False
End Sub

Private Sub Comm()
On Error GoTo Err
    If MSComm2.PortOpen = True Then
        MSComm2.PortOpen = False
    Else
        MSComm2.CommPort = PrtNumb
        MSComm2.Settings = CONF
        MSComm2.PortOpen = True
    End If
Exit Sub
Err:
    MsgBox "Error opening Comm Port #" & PrtNumb, vbCritical, "Comm Port Error"
End Sub

Private Sub Timer3_Timer()
If BVD = 6 Then
        If CNT > 0 Then
            Label2 = Round((CNT / 6), 2)
        Else
            Label2 = "0.00"
        End If
    Call Comm
    Timer4.Enabled = False
    RunTest.BackColor = &H8000000F
    RunTest.Caption = "Run Test"
    CNT = 0
    BVD = 0
    Timer3.Enabled = False
    Done.Enabled = True
End If
BVD = BVD + 1
End Sub

Private Sub Timer4_Timer()
If RunTest.BackColor = &H40& Then
    RunTest.BackColor = vbRed
Else
    RunTest.BackColor = &H40&
End If
End Sub
'                             Code by 0x34 - 2004
