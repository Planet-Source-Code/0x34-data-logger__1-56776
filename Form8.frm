VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peak Detect Trigger Settings"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr1 
      Caption         =   "Alert Sound"
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
      Begin VB.CommandButton Command4 
         Caption         =   "Accept"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   2160
         Picture         =   "Form8.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   375
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   240
         Pattern         =   "*.wav"
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Beep only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   260
         TabIndex        =   7
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label SelSnd 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SelSnd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Selected Sound"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
      _Version        =   327681
      Max             =   255
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Max             =   255
   End
   Begin VB.Label Label4 
      Caption         =   "0x34"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Trigger Level Adjust"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   800
      Width           =   735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempSnd As String

Private Sub Check1_Click()
If Check1 Then
    BpOnly = True
    File1.Enabled = False
    SelSnd.Enabled = False
    Label2.Enabled = False
    Command4.Enabled = False
Else
    BpOnly = False
    Label2.Enabled = True
    File1.Enabled = True
    SelSnd.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Check1 Then
        Beep
    Else
        CLIK = sndPlaySound(TempSnd, 1)
    End If
End Sub

Private Sub Command4_Click()
    SoundB = TempSnd
    SelSnd = SoundB
End Sub

Private Sub File1_Click()
    TempSnd = "C:\Windows\Media\" & File1
End Sub

Private Sub File1_DblClick()
    TempSnd = "C:\Windows\Media\" & File1
    CLIK = sndPlaySound(TempSnd, 1)
End Sub

Public Sub Form_Load()
File1.Path = "C:\Windows\Media"
UpDown1 = TrLev
ProgressBar1 = TrLev
Label1 = TrLev
SelSnd = SoundB
TempSnd = SoundB
If BpOnly Then
    Check1 = 1
    File1.Enabled = False
    SelSnd.Enabled = False
    Label2.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
Else
    Check1 = 0
    Label2.Enabled = True
    File1.Enabled = True
    SelSnd.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open REXsav For Output As #1
        Print #1, SoundB
        If BpOnly Then
            Print #1, 1
        Else
            Print #1, 0
        End If
            Print #1, TrLev
    Close #1
End Sub

Public Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BX As Variant
    If X > 999 Then
        BX = Left$(X, 3)
        GoTo RWEQ
    End If
    If X > 99 Then
        BX = Left$(X, 2)
        GoTo RWEQ
    End If
    BX = Left$(X, 1)
RWEQ:
    UpDown1 = BX
End Sub

Private Sub ProgressBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BX As Variant
If Button = 1 Then
    If X > 999 Then
        BX = Left$(X, 3)
        GoTo RWEQ
    End If
    If X > 99 Then
        BX = Left$(X, 2)
        GoTo RWEQ
    End If
    BX = Left$(X, 1)
RWEQ:
    If BX < 256 Then
        UpDown1 = BX
    End If
End If
End Sub

Private Sub SelSnd_Click()
TempSnd = SoundB
CLIK = sndPlaySound(TempSnd, 1)
End Sub

Public Sub UpDown1_Change()
TrLev = UpDown1
Call SetA
End Sub

Public Sub SetA()
On Error GoTo ERRORT
    ProgressBar1 = TrLev
    Label1 = TrLev
ERRORT:
    Exit Sub
End Sub
