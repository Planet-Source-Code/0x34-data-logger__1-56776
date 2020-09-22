VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust Colors"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command16 
      Caption         =   "Save As Custom"
      Height          =   495
      Left            =   3960
      TabIndex        =   28
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Custom"
      Height          =   255
      Left            =   5400
      TabIndex        =   27
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Load Default"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Right Cursor"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Left Cursor"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Apply to Zero Bar"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Pre3 
      Caption         =   "PreSet3"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Pre2 
      Caption         =   "PreSet 2"
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Pre1 
      Caption         =   "PreSet 1"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Background"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apply to Graph"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply to Grid"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DONE"
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Max             =   255
      TickStyle       =   2
      TickFrequency   =   255
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Max             =   255
      TickStyle       =   2
      TickFrequency   =   255
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Max             =   255
      TickStyle       =   2
      TickFrequency   =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0x34"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   26
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "BLUE"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "GREEN"
      Height          =   255
      Left            =   -120
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RED"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fr As Integer
Dim Fg As Integer
Dim Fb As Integer

Public Sub Command1_Click()
Form2.Visible = False
Form1.Picture1.Refresh
Call Form1.PosAdju
Form1.ReDrCur
Unload Me
End Sub

Private Sub Command10_Click()
CZr = Fr
CZg = Fg
CZb = Fb
Form1.Picture1.Refresh
Call Form1.Graph
Command9.BackColor = RGB(CZr, CZg, CZb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command11_Click()
CLRx = Fr
CLGx = Fg
CLBx = Fb
Form1.Picture1.Refresh
Call Form1.Graph
Command13.BackColor = RGB(CLRx, CLGx, CLBx)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command12_Click()
CRRx = Fr
CRGx = Fg
CRBx = Fb
Form1.Picture1.Refresh
Call Form1.Graph
Command14.BackColor = RGB(CRRx, CRGx, CRBx)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command13_Click()
Slider1.Value = CLRx
Slider2.Value = CLGx
Slider3.Value = CLBx
Fr = CLRx
Fg = CLGx
Fb = CLBx
Label7.ForeColor = RGB(CLRx, CLGx, CLBx)
Call UpDateC
End Sub

Private Sub Command14_Click()
Slider1.Value = CRRx
Slider2.Value = CRGx
Slider3.Value = CRBx
Fr = CRRx
Fg = CRGx
Fb = CRBx
Label7.ForeColor = RGB(CRRx, CRGx, CRBx)
Call UpDateC
End Sub

Private Sub Command15_Click()
Call ColLoad
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Slider1.Value = Fr
Slider2.Value = Fg
Slider3.Value = Fb
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
Form1.Picture1.Refresh
Form1.Picture1.BackColor = RGB(CBR, CBg, CBb)
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command16_Click()
Call ColSav
End Sub

Public Sub Command2_Click()
Cr = Slider1.Value
Cg = Slider2.Value
Cb = Slider3.Value
Command2.BackColor = RGB(Cr, Cg, Cb)
Form1.Picture1.Refresh
Call Form1.Graph
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Public Sub Command3_Click()
Gr = Slider1.Value
Gg = Slider2.Value
Gb = Slider3.Value
Command3.BackColor = RGB(Gr, Gg, Gb)
Form1.Picture1.Refresh
Call Form1.Graph
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command4_Click()
CBR = Fr
CBg = Fg
CBb = Fb
Form1.Picture1.BackColor = RGB(CBR, CBg, CBb)
Form1.Picture1.Refresh
Call Form1.Graph
Command8.BackColor = RGB(CBR, CBg, CBb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command5_Click()
Slider1.Value = Cr
Slider2.Value = Cg
Slider3.Value = Cb
Fr = Cr
Fg = Cg
Fb = Cb
Label7.ForeColor = RGB(Cr, Cg, Cb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command6_Click()
Slider1.Value = Gr
Slider2.Value = Gg
Slider3.Value = Gb
Fr = Gr
Fg = Gg
Fb = Gb
Label7.ForeColor = RGB(Gr, Gg, Gb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command7_Click()
Cr = 0
Cg = 50
Cb = 100
Gr = 0
Gg = 255
Gb = 0
CLRx = 0
CLGx = 255
CLBx = 255
CRRx = 255
CRGx = 255
CRBx = 0
CZr = 255
CZg = 0
CZb = 0
CBR = 0
CBg = 0
CBb = 0
Form1.Picture1.BackColor = vbBlack
Form1.Picture1.Refresh
Call Form1.Graph
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
Command13.BackColor = RGB(CLRx, CLGx, CLBx)
Command14.BackColor = RGB(CRRx, CRGx, CRBx)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command8_Click()
Slider1.Value = CBR
Slider2.Value = CBg
Slider3.Value = CBb
Fr = CBR
Fg = CBg
Fb = CBb
Label7.ForeColor = RGB(CBR, CBg, CBb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Command9_Click()
Slider1.Value = CZr
Slider2.Value = CZg
Slider3.Value = CZb
Fr = CZr
Fg = CZg
Fb = CZb
Label7.ForeColor = RGB(CZr, CZg, CZb)
Call UpDateC
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Public Sub Form_Load()
Dim FART As String
    FART = Chr(&H37 - 7)
    FART = FART & Chr(&H7B - 3)
    FART = FART & Chr(&H39 - 6)
    FART = FART & Chr(&H38 - 4)
    Label7 = FART
    FART = ""
Fr = 193
Fg = 193
Fb = 193
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Slider1.Value = Fr
Slider2.Value = Fg
Slider3.Value = Fb
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
Command13.BackColor = RGB(CLRx, CLGx, CLBx)
Command14.BackColor = RGB(CRRx, CRGx, CRBx)
Label6 = Fb
Label5 = Fg
Label4 = Fr
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command1_Click
End Sub

Private Sub Pre1_Click()
Cr = 90
Cg = 90
Cb = 90
Gr = 255
Gg = 255
Gb = 0
CZr = 255
CZg = 0
CZb = 0
CBR = 0
CBg = 0
CBb = 0
Fr = 193
Fg = 193
Fb = 193
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Slider1.Value = Fr
Slider2.Value = Fg
Slider3.Value = Fb
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
Form1.Picture1.Refresh
Form1.Picture1.BackColor = RGB(CBR, CBg, CBb)
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Pre2_Click()
Cr = 193
Cg = 193
Cb = 193
Gr = 0
Gg = 0
Gb = 0
CZr = 255
CZg = 0
CZb = 0
CBR = 255
CBg = 255
CBb = 255
Fr = 193
Fg = 193
Fb = 193
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Slider1.Value = Fr
Slider2.Value = Fg
Slider3.Value = Fb
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
Form1.Picture1.Refresh
Form1.Picture1.BackColor = RGB(CBR, CBg, CBb)
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Pre3_Click()
Cr = 62
Cg = 255
Cb = 108
Gr = 255
Gg = 255
Gb = 255
CZr = 255
CZg = 255
CZb = 0
CBR = 62
CBg = 81
CBb = 108
Fr = 193
Fg = 193
Fb = 193
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Slider1.Value = Fr
Slider2.Value = Fg
Slider3.Value = Fb
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
Form1.Picture1.Refresh
Form1.Picture1.BackColor = RGB(CBR, CBg, CBb)
Call Form1.PosAdju
Form1.ReDrCur
End Sub

Private Sub Slider1_Scroll()
Fr = Slider1
Label7.ForeColor = RGB(Fr, Fg, Fb)
Label4 = Fr
Call UpDateC
End Sub

Private Sub Slider2_Scroll()
Fg = Slider2
Label7.ForeColor = RGB(Fr, Fg, Fb)
Label5 = Fg
Call UpDateC
End Sub

Public Sub UpDateC()
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Label6 = Fb
Label5 = Fg
Label4 = Fr
Command5.BackColor = RGB(Cr, Cg, Cb)
Command6.BackColor = RGB(Gr, Gg, Gb)
Slider1.Value = Fr
Slider2.Value = Fg
Slider3.Value = Fb
Label7.ForeColor = RGB(Fr, Fg, Fb)
Command2.BackColor = RGB(Cr, Cg, Cb)
Command3.BackColor = RGB(Gr, Gg, Gb)
Command9.BackColor = RGB(CZr, CZg, CZb)
Command8.BackColor = RGB(CBR, CBg, CBb)
'Form1.ReDrCur
End Sub

Private Sub Slider3_Scroll()
Fb = Slider3
Label7.ForeColor = RGB(Fr, Fg, Fb)
Label6 = Fb
Call UpDateC
End Sub
'                             Code by 0x34 - 2004
