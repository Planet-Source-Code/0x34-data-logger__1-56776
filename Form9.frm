VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DeBug"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2085
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   2085
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   735
      Left            =   240
      Picture         =   "Form9.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
Label1(0) = "RUUN = " & RUUN
Label1(1) = "ASH = " & ASH
Label1(2) = "RecSTP = " & RecStp
Label1(3) = "Joker = " & Joker
Label1(4) = "SavD = " & SavD
Label1(5) = "Zero = " & Zero
Label1(6) = "TEMP = " & TEMP
Label1(7) = "DataChunk = " & DataChunck
Label1(8) = "Left Cur = " & Posi1
Label1(9) = "Right Cur = " & Posi2
Label1(10) = "LCLoc = " & LCLoc
Label1(11) = "RCLoc = " & RCLoc
Label1(12) = "SB1 = " & Form1.SB1
Label1(13) = "ZSP = " & ZSP
Label1(14) = "RecLocLeft = " & (ZSP + LCLoc)
Label1(15) = "RecLocRight = " & (ZSP + RCLoc)
End Sub
