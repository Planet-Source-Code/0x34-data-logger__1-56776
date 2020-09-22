VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You cannot HEX EDIT this stuff you wannabe programmer ass hole!"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TSp 
      Enabled         =   0   'False
      Interval        =   90
      Left            =   6000
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   1095
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   4575
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "NON-Hex Editable Credits ass hole!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "No Code Stealing allowed!"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "You cannot edit this and make it look like you wrote it!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   450
         Left            =   120
         TabIndex        =   5
         Top             =   165
         Width           =   4335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2120
      Left            =   120
      Picture         =   "Form4.frx":08CA
      ScaleHeight     =   2055
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   90
      Width           =   1880
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   16
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   15
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   14
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   13
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   12
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   11
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   10
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   9
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   8
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   7
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   6
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   5
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   4
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   3
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   2
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   1
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   0
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "...my patient Wife"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "Ox34's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Keep Looking Dip Shit!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'  NOTE**  The vulgar text inserted into the labels of this form is ment
'          to be located by those little kids who like to alter the
'          programmer's credits with a HEX editor and try to pass the
'          program off as their own.  All the text will be replaced
'          during runtime with encoded HEX text.  The code on this
'          page fills the labels.
'
'                                                             0x34

Dim Cycle As Integer
Private Sub Form_Load()
Dim FART As String
Cycle = 0
    FART = Chr((&H33 - 3))
    FART = FART & Chr((&H7B - 3))
    FART = FART & Chr((&H36 - 3))
    FART = FART & Chr((&H37 - 3))
    FART = FART & Chr((&H2A - 3))
    FART = FART & Chr((&H76 - 3))
    Label4 = FART
FART = ""
    FART = FART & Chr((&H46 - 3))
    FART = FART & Chr((&H77 - 8))
    FART = FART & Chr((&H67 - 3))
    FART = FART & Chr((&H6C - 3))
    FART = FART & Chr((&H71 - 3))
    FART = FART & Chr((&H6A - 3))
    FART = FART & Chr((&H23 - 3))
    FART = FART & Chr((&H45 - 3))
    FART = FART & Chr((&H7C - 3))
    FART = FART & Chr((&H23 - 3))
    FART = FART & Chr((&H36 - 6))
    FART = FART & Chr((&H7B - 3))
    FART = FART & Chr((&H39 - 6))
    FART = FART & Chr((&H37 - 3))
    FART = FART & Chr((&H23 - 3))
    FART = FART & Chr((&H30 - 3))
    FART = FART & Chr((&H23 - 3))
    FART = FART & Chr((&H35 - 3))
    FART = FART & Chr((&H36 - 6))
    FART = FART & Chr((&H39 - 9))
    FART = FART & Chr((&H37 - 3))
    Label7 = FART
    FART = ""
FART = FART & Chr((&H47 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H4F - 3))
FART = FART & Chr((&H72 - 3))
FART = FART & Chr((&H6A - 3))
FART = FART & Chr((&H6A - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H75 - 3))
Label1 = FART
FART = ""
FART = FART & Chr((&H55 - 3))
FART = FART & Chr((&H56 - 3))
FART = FART & Chr((&H35 - 3))
FART = FART & Chr((&H36 - 3))
FART = FART & Chr((&H35 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H47 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H4F - 3))
FART = FART & Chr((&H72 - 3))
FART = FART & Chr((&H6A - 3))
FART = FART & Chr((&H6A - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H75 - 3))
Label6 = FART
FART = ""
FART = FART & Chr((&H47 - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H79 - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H6F - 3))
FART = FART & Chr((&H72 - 3))
FART = FART & Chr((&H73 - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H67 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H69 - 3))
FART = FART & Chr((&H72 - 3))
FART = FART & Chr((&H75 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H6B - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H46 - 3))
FART = FART & Chr((&H49 - 3))
FART = FART & Chr((&H36 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H53 - 3))
FART = FART & Chr((&H75 - 3))
FART = FART & Chr((&H72 - 3))
FART = FART & Chr((&H6D - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H66 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H31 - 3))
Label8 = FART
FART = ""
FART = FART & Chr((&H31 - 3))
FART = FART & Chr((&H31 - 3))
FART = FART & Chr((&H31 - 3))
FART = FART & Chr((&H70 - 3))
FART = FART & Chr((&H7C - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H73 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H6C - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H71 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H5A - 3))
FART = FART & Chr((&H6C - 3))
FART = FART & Chr((&H69 - 3))
FART = FART & Chr((&H68 - 3))
Label5 = FART
FART = ""
FART = FART & Chr((&H45 - 4))
FART = FART & Chr((&H66 - 4))
FART = FART & Chr((&H75 - 6))
FART = FART & Chr((&H78 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H23 - 3))
FART = FART & Chr((&H36 - 6))
FART = FART & Chr((&H7B - 3))
FART = FART & Chr((&H36 - 3))
FART = FART & Chr((&H37 - 3))
FART = FART & Chr((&H2A - 3))
FART = FART & Chr((&H77 - 4))
FART = FART & Chr((&H26 - 6))
FART = FART & Chr((&H47 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H77 - 3))
FART = FART & Chr((&H64 - 3))
FART = FART & Chr((&H4F - 3))
FART = FART & Chr((&H72 - 3))
FART = FART & Chr((&H6A - 3))
FART = FART & Chr((&H6A - 3))
FART = FART & Chr((&H68 - 3))
FART = FART & Chr((&H75 - 3))
Form4.Caption = FART
FART = ""
End Sub

Private Sub Form_LostFocus()
Form4.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
TSp.Enabled = False
End Sub

Private Sub Picture1_Click()
If TSp.Enabled = False Then
    TSp.Enabled = True
Else
    TSp.Enabled = False
    For t = 0 To 16
        Shape1(t).Visible = False
    Next t
End If
End Sub

Private Sub TSp_Timer()
Cycle = Cycle + 1
If Cycle = 17 Then
    Cycle = 0
    Shape1(Cycle).Visible = True
    Shape1(16).Visible = False
Else
    Shape1(Cycle).Visible = True
    Shape1(Cycle - 1).Visible = False
End If

End Sub
'                             Code by 0x34 - 2004
