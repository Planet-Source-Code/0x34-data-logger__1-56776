VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comm Port Settings"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2475
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer9 
      Interval        =   10
      Left            =   2280
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Default"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Settings"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   735
      Left            =   360
      Picture         =   "Form3.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Comm Port"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Stop Bits"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Legnth"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Parity"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Baud Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Altered As Boolean
Dim Ai As Integer
Dim Bi As Integer

Private Sub Command1_Click()
BRate = Text1
Parity = Text2
NBytes = Text3
StopBits = Text4
If Text5 > 0 And Text5 < 5 Then
    PrtNumb = Text5
End If
CONF = BRate & "," & Parity & "," & NBytes & "," & StopBits
Timer9.Enabled = False
Unload Me
End Sub

Private Sub Command2_Click()
    Open CommSett For Output As #1
        Print #1, Text1
        Print #1, Text2
        Print #1, Text3
        Print #1, Text4
        If Text5 > 0 And Text5 < 5 Then
            Print #1, Text5
        Else
            Print #1, 0
        End If
    Close #1
    Command2.BackColor = &H8000000F
    Altered = False
End Sub

Private Sub Command3_Click()
BRate = "9600"
Parity = "N"
StopBits = "1"
NBytes = "8"
PrtNumb = 1
Text1 = BRate
Text2 = Parity
Text3 = NBytes
Text4 = StopBits
Text5 = PrtNumb
CONF = BRate & "," & Parity & "," & NBytes & "," & StopBits
End Sub

Private Sub Form_Load()
On Error GoTo Yelling
    Open CommSett For Input As #1
        Input #1, BRate
        Input #1, Parity
        Input #1, NBytes
        Input #1, StopBits
        Input #1, PrtNumb
    Close #1
Text1 = BRate
Text2 = Parity
Text3 = NBytes
Text4 = StopBits
Text5 = PrtNumb
    CONF = BRate & "," & Parity & "," & NBytes & "," & StopBits
    Altered = False
    Timer9.Enabled = True
    Ai = 0
    Bi = 0
Exit Sub
Yelling:
    BRate = "9600"
    Parity = "N"
    StopBits = "1"
    NBytes = "8"
    PrtNumb = 1
    Text1 = BRate
    Text2 = Parity
    Text3 = NBytes
    Text4 = StopBits
    Text5 = PrtNumb
    CONF = BRate & "," & Parity & "," & NBytes & "," & StopBits
    Altered = False
    Timer9.Enabled = True
End Sub

Private Sub Text1_Change()
Altered = True
Command2.BackColor = vbRed
End Sub

Private Sub Text2_Change()
Altered = True
Command2.BackColor = vbRed
End Sub

Private Sub Text3_Change()
Altered = True
Command2.BackColor = vbRed
End Sub

Private Sub Text4_Change()
Altered = True
Command2.BackColor = vbRed
End Sub

Private Sub Text5_Change()
Altered = True
Command2.BackColor = vbRed
End Sub

Private Sub Timer9_Timer()
Ai = Ai + 1
If Ai < 3 Then
    Altered = False
    Command2.BackColor = &H8000000F
End If
If Altered Then
    If Bi > 50 Then
        If Command2.BackColor = vbRed Then
            Command2.BackColor = &H8000000F
        Else
            Command2.BackColor = vbRed
        End If
        Bi = 0
    Else
        Bi = Bi + 1
    End If
End If
End Sub
'                             Code by 0x34 - 2004
