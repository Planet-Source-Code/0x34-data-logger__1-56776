VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graph Preferences"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Mirror Graph"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Marker Size"
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin MSComCtl2.FlatScrollBar FSB1 
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         LargeChange     =   10
         Min             =   612
         Max             =   12
         Orientation     =   1245184
         SmallChange     =   10
         Value           =   20
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   480
      Top             =   480
   End
   Begin VB.PictureBox P1 
      Height          =   600
      Left            =   240
      ScaleHeight     =   600
      ScaleMode       =   0  'User
      ScaleWidth      =   1000
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
   Begin MSComCtl2.FlatScrollBar VScroll1 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      Min             =   20
      Max             =   1
      Orientation     =   1245184
      Value           =   20
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Line Weight"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DrWdt < 20 Then
    DrWdt = DrWdt + 1
End If
Label1 = DrWdt
Call UpDat
End Sub

Private Sub Command2_Click()
Label1 = DrWdt
Call UpDat
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
    If Check1 = 1 Then
        Mirror = True
    Else
        Mirror = False
    End If
Call Form1.PosAdju
Form1.ReDrCur
Unload Me
End Sub

Private Sub FSB1_change()
Label2 = (FSB1 - 12)
Marker = FSB1
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Form_Load()
FSB1 = Marker
Label2 = (Marker - 12)
VScroll1 = DrWdt
P1.BackColor = vbBlack
P1.DrawWidth = DrWdt
P1.Line (100, 300)-(900, 300), RGB(0, 255, 0)
Call UpDat
Timer1.Enabled = True
Label1 = DrWdt
End Sub

Private Sub UpDat()
P1.BackColor = vbBlack
P1.DrawWidth = VScroll1
P1.Line (100, 300)-(900, 300), RGB(0, 255, 0)
    If Mirror Then
        Check1 = 1
    Else
        Check1 = 0
    End If
End Sub

Private Sub Timer1_Timer()
Call UpDat
Timer1.Enabled = False
End Sub

Private Sub VScroll1_Change()
Label1 = VScroll1
DrWdt = VScroll1
Call UpDat
Call Form1.PosAdju
Form1.ReDrCur
End Sub
'                             Code by 0x34 - 2004
