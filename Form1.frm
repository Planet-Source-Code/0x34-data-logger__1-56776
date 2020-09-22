VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider Slider1 
      Height          =   5895
      Left            =   10200
      TabIndex        =   38
      ToolTipText     =   "Zero Value Set"
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   10398
      _Version        =   393216
      Orientation     =   1
      Max             =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ZOOM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   135
      Left            =   7680
      TabIndex        =   36
      ToolTipText     =   "De-Bug"
      Top             =   7080
      Width           =   135
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   34
      ToolTipText     =   "Recording Length Bar"
      Top             =   8100
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.FlatScrollBar CBL 
      Height          =   255
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Left Cursor"
      Top             =   6000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Max             =   1000
      Orientation     =   1245185
   End
   Begin MSComCtl2.FlatScrollBar CBR 
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      ToolTipText     =   "Right Cursor"
      Top             =   6000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Max             =   1000
      Orientation     =   1245185
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cursors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      TabIndex        =   20
      Top             =   6720
      Width           =   2535
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Hz"
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Sec"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Label LabFreq 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label11"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   960
         Width           =   615
      End
      Begin VB.Label LabDIFF 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100.10"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Difference"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LabAmpR 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LabAmpL 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Frequency"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
   End
   Begin MSComCtl2.FlatScrollBar SB1 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6360
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   12
      Top             =   7200
      Width           =   2415
      Begin VB.CommandButton Command7 
         Caption         =   "Reset"
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton PAUSE 
         Height          =   600
         Left            =   840
         Picture         =   "Form1.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   172
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Run Once"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   10
      Top             =   6720
      Width           =   975
      Begin VB.CommandButton Command2 
         Caption         =   "Peak"
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin MSComCtl2.UpDown Mag1 
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   873
         _Version        =   393216
         Value           =   10
         Increment       =   10
         Max             =   150
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   440
         TabIndex        =   11
         Top             =   370
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Horiz Sweep Adjust"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   2655
      Begin MSComctlLib.Slider Slider2 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         SelStart        =   3
         Value           =   3
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Per Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   880
         Width           =   975
      End
   End
   Begin VB.CheckBox Check1 
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
      Left            =   10200
      TabIndex        =   2
      ToolTipText     =   "Grid"
      Top             =   6360
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Zero ADJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1296
         _Version        =   393216
         Value           =   3000
         Increment       =   300
         Max             =   0
         Min             =   6000
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Z"
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9240
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      DragIcon        =   "Form1.frx":0D0C
      Height          =   6000
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   6000
      ScaleMode       =   0  'User
      ScaleWidth      =   1000
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   9000
         Top             =   1920
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   8400
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "dlg"
         FontSize        =   12
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   9120
         Top             =   1200
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   8400
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         Handshaking     =   2
         InputLen        =   1
         RThreshold      =   1
         SThreshold      =   1
         EOFEnable       =   -1  'True
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label15"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH.mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9480
      TabIndex        =   42
      Top             =   8100
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label14"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yy MM dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8520
      TabIndex        =   41
      Top             =   8100
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   15
      Left            =   10080
      TabIndex        =   40
      Top             =   8160
      Width           =   15
   End
   Begin VB.Label Label12 
      Caption         =   "255"
      Height          =   255
      Left            =   10200
      TabIndex        =   39
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Sec."
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
      Left            =   7380
      TabIndex        =   9
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Elapsed"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   6840
      Width           =   735
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MNULoad 
         Caption         =   "Load file"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save file"
      End
      Begin VB.Menu VCxxz 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu MnuSp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuABOUT 
         Caption         =   "About"
      End
      Begin VB.Menu MnuSp2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEXIT 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "Tools"
      Begin VB.Menu MnuTrigr 
         Caption         =   "Trigger Level"
      End
      Begin VB.Menu mnbvcxz 
         Caption         =   "-"
      End
      Begin VB.Menu MNUPing 
         Caption         =   "Ping Test"
      End
      Begin VB.Menu mnuspx 
         Caption         =   "-"
      End
      Begin VB.Menu MNUclock 
         Caption         =   "Clock Select"
      End
   End
   Begin VB.Menu mnuPrefs 
      Caption         =   "Preferences"
      Begin VB.Menu MnuColor 
         Caption         =   "Colors"
      End
      Begin VB.Menu MnuGraph 
         Caption         =   "Graph"
      End
      Begin VB.Menu mnueera 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSN 
         Caption         =   "Save Nag"
         Checked         =   -1  'True
      End
      Begin VB.Menu jlkklj 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCOMM 
         Caption         =   "COMM Config"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'             **************************************************
'             **************************************************
'             ********    0x34's Data Logger - 2004     ********
'             ********   For Embedded system design.    ********
'             ********  Records up to 5 hours of RS232  ********
'             ********     Data in Graphical form       ********
'             ********          like a DSO.             ********
'             ********  Coded out of sheer nessessity.  ********
'             **************************************************
'             **************************************************
'             ********  Time & Date added / 10 mS added ********
'             **************************************************
'             **************************************************
'
Dim Yr As Long
Dim Xr As Long
Dim Yn As Long
Dim Xn As Long
Dim Grad As Boolean
Dim StepRate As Integer
Dim TMR As Currency
Dim ONCE As Boolean
Dim ModStr As String
Dim Command As String
Dim FUK As String ' <---------  That's right... FUK!
Dim PORTC As Boolean
Dim MAG As Integer
Dim MBV As Boolean
Dim Yx As Long
Dim NB As Long
'********************************************
Private Sub CBL_Change() ' Cursor Bar LEFT
If CBL > (CBR - 1) Then Exit Sub
    If Zoom = False And ZEFin = False Then
        LOrPos = CBL
        Posi1 = (DataChunck + CBL)
        LCLoc = ((Posi1 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
If Zoom Then
    LabAmpL = DATA(ZSP + CBL)
    LabAmpR = DATA(ZSP + CBR)
End If
Call PlaceCur
Call Calculate
End Sub
'********************************************
Private Sub CBL_Scroll() ' Cursor Bar LEFT
If CBL > (CBR - 1) Then Exit Sub
    If Zoom = False And ZEFin = False Then
        LOrPos = CBL
        Posi1 = (DataChunck + CBL)
        LCLoc = ((Posi1 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
If Zoom Then
    LabAmpL = DATA(ZSP + CBL)
    LabAmpR = DATA(ZSP + CBR)
End If
Call PlaceCur
Call Calculate
End Sub
'********************************************
Private Sub CBR_Change() ' Cursor Bar RIGHT
If CBR < (CBL + 1) Then Exit Sub
    If Zoom = False And ZEFin = False Then
        ROrPos = CBR
        Posi2 = (DataChunck + CBR)
        RCLoc = ((Posi2 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
    If Zoom = False And ZEFin = False Then ROrPos = CBR
If Zoom Then
    LabAmpL = DATA(ZSP + CBL)
    LabAmpR = DATA(ZSP + CBR)
End If
Call PlaceCur
Call Calculate
End Sub
'********************************************
Private Sub CBR_Scroll() ' Cursor Bar RIGHT
If CBR < (CBL + 1) Then Exit Sub
    If Zoom = False And ZEFin = False Then
        ROrPos = CBR
        Posi2 = (DataChunck + CBR)
        RCLoc = ((Posi2 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
If Zoom Then
    LabAmpL = DATA(ZSP + CBL)
    LabAmpR = DATA(ZSP + CBR)
End If
Call PlaceCur
Call Calculate
End Sub
'********************************************
Private Sub Calculate() ' Cursor Calculation Function
On Error GoTo YRT
Dim Aq As Single
Dim Bq As Single
    Aq = 10 * (0.001 * (CBR - CBL))
    Bq = Aq
        If Aq < 1 Then
            Aq = Aq * 1000
            Label10 = "mS"
        Else
            Label10 = "Sec"
        End If
    LabDIFF = Aq
    LabFreq = Round(1 / Bq, 2)
Exit Sub
YRT:
    Exit Sub
End Sub
'********************************************
Private Sub Check1_Click() ' Grid Select CheckBox
If Check1 Then
    If Picture1.ScaleWidth < 29999 Then
        CBL.Enabled = True
        CBR.Enabled = True
        If Slider2 > 0 Or Slider2 < 6 Then
            If Label7 = "50 mS " Then
                Command5.Enabled = False
            Else
                Command5.Enabled = True
            End If
        Else
            Command5.Enabled = False
        End If
        LabAmpL = "0.0"
        LabAmpR = "0.0"
        LabDIFF = "0.00"
        LabFreq = "0.00"
        LabAmpL.Enabled = True
        LabAmpR.Enabled = True
        LabFreq.Enabled = True
        LabDIFF.Enabled = True
        Call Calculate
        Cursors = True
    Else
        Cursors = False
    End If
Else
    CBL.Enabled = False
    CBR.Enabled = False
    Command5.Enabled = False
    LabAmpL = "0.0"
    LabAmpR = "0.0"
    LabDIFF = "0.00"
    LabFreq = "0.00"
    LabAmpL.Enabled = False
    LabAmpR.Enabled = False
    LabFreq.Enabled = False
    LabDIFF.Enabled = False
    Cursors = False
End If
Call Graph
If RecStp > 0 Then
    Call PosAdju
End If
Call ReDrCur
End Sub
'********************************************
Private Sub Command1_Click() ' DeBug Screen
If Form9.Visible Then
    Unload Form9
Else
    Form9.Show
End If
End Sub
'********************************************
Private Sub Command2_Click() ' Peak
If Form8.Visible Then Unload Form8
If TrigX Then
    TrigX = False
    Command2.BackColor = &H8000000F
Else
    TrigX = True
    Command2.BackColor = vbGreen
End If
End Sub
'********************************************
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Form8.Visible Then
        Unload Form8
    Else
        Form8.Show
    End If
End If
End Sub
'********************************************
Private Sub Command3_Click() ' Run Once
If Form8.Visible Then Unload Form8
Call RENEW
    ONCE = True
    TMR = 0
Call PSIM
RecTime = Format(Time, "Medium time")
RecDate = Format(Date, "short Date")
End Sub
'********************************************
Private Sub PSIM() ' Begin Single Page Sample
    If RUUN = True Then
        RUUN = False
        Command3.Enabled = True
        Command3.BackColor = &H8000000F
        Call Comm
        MnuCOMM.Enabled = True
        MNUPing.Enabled = True
        Slider2.Enabled = True
        CBL.Enabled = True
        CBR.Enabled = True
    Else
        If RecStp > 0 Then
        Call SnvNag
        End If
        If Joker Then Exit Sub
        Call VOIDd
        RECORD = True
        Call Comm
        If MSComm1.PortOpen = False Then Exit Sub
        Frame = 0
        SB1.Max = Frame
        RecStp = 0
        RUUN = True
        Call RENEW
        Command3.Enabled = False
            If ONCE Then
                Command3.BackColor = RGB(255, 255, 0)
            End If
        MnuCOMM.Enabled = False
        MNUPing.Enabled = False
        Slider2.Enabled = False
        CBL.Enabled = False
        CBR.Enabled = False
    End If
Call GoSet
If RUUN = False Then
    Call ReDrCur
End If
End Sub
'********************************************
Private Sub Comm() ' RS232 Comm Port Adjust and Control
On Error GoTo Err
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    Else
        MSComm1.CommPort = PrtNumb
        MSComm1.Settings = CONF
        MSComm1.PortOpen = True
    End If
Exit Sub
Err:
    MsgBox "Error opening Comm Port #" & PrtNumb, vbCritical, "Comm Port Error"
    RUUN = False
End Sub
'********************************************
Private Sub Command4_Click() ' Zero Adjust Down
If Zero > 5985 Then Exit Sub
    Zero = Zero + 10
    Picture1.Refresh
Call Graph
If RecStp > 0 Then
    Call PosAdju
End If
Call ReDrCur
    UpDown1.Value = Zero
End Sub
'********************************************
Private Sub Command5_Click() 'Zoom
If Zoom Then
    Picture1.Cls
    Zoom = False
    MNUclock.Enabled = True
    Call Slider2_Scroll
    Command5.BackColor = &H8000000F
    Slider2.Enabled = True
    SB1.Enabled = True
    Command3.Enabled = True
    PAUSE.Enabled = True
    Command7.Enabled = True
    Check1.Enabled = True
    CBR = ROrPos
    CBL = LOrPos
    ZEFin = False
Else
    Zoom = True
    MNUclock.Enabled = False
    ZEFin = True
    Picture1.Cls
    Check1.Enabled = False
    Slider2.Enabled = False
    SB1.Enabled = False
    Command3.Enabled = False
    PAUSE.Enabled = False
    Command7.Enabled = False
    ZSP = LCLoc
    Command5.BackColor = vbGreen
    MousePointer = vbHourglass
    Call TZAdj
    DataChunck = DataChunck + Posi1
    MousePointer = vbDefault
    Call ReDrCur
    On Error GoTo ExSbER
    If Check1 And Cursors And RUUN = False Then
        If Button = 1 Then
            If X > (CBR - 1) Then Exit Sub
            If X < 1 Then Exit Sub
            CBL = X
            Posi1 = (DataChunck + X)
            LCLoc = ((Posi1 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
        End If
        If Button = 2 Then
            If X < (CBL + 1) Then Exit Sub
            If X > Picture1.ScaleWidth Then Exit Sub
                CBR = X
                Posi2 = (DataChunck + X)
            End If
        End If
    End If
ExSbER:
    Exit Sub
End Sub
'********************************************
Private Sub Command7_Click() ' Reset
If Form8.Visible Then Unload Form8
Call RSETa
End Sub
'********************************************
Public Sub RSETa() ' Main Reset Function
Dim NV As Integer
If RecStp > 1 And SavD = False Then
    NV = MsgBox("Do you want to save the current recording? ", vbYesNoCancel, "Save Sample?")
    If NV = vbYes Then
        Call mnuSave_Click
        NV = MsgBox("Confirm Clear all Data?  ", vbYesNo, "Reset")
        If NV = vbNo Then
            Exit Sub
        End If
    End If
    If NV = vbCancel Then
        Exit Sub
    End If
End If
Timer3.Enabled = True
If Cursors Then
    If Check1 Then
        CBL.Enabled = True
        CBR.Enabled = True
        CBL = 2
        CBR = (Picture1.ScaleWidth - 2)
        LabAmpL = "0.0"
        LabAmpR = "0.0"
        LabDIFF = "0.00"
        LabFreq = "0.00"
        LabAmpL.Enabled = True
        LabAmpR.Enabled = True
        LabFreq.Enabled = True
        LabDIFF.Enabled = True
        Cursors = True
        Call Calculate
    Else
        CBL.Enabled = False
        CBR.Enabled = False
        LabAmpL = "0.0"
        LabAmpR = "0.0"
        LabDIFF = "0.00"
        LabFreq = "0.00"
        LabAmpL.Enabled = False
        LabAmpR.Enabled = False
        LabFreq.Enabled = False
        LabDIFF.Enabled = False
        Cursors = False
    End If
End If
    If MnuSN.Checked = True Then
        SavD = False
    Else
        SavD = True
    End If
Call RENEW
RecStp = 0
PB1 = RecStp
RECORD = False
Xni = 0
Yni = Zero
ONCE = False
Call VOIDd
Joker = False
    Frame = 0
    SB1.Max = Frame
    TMR = 0
        RUUN = False
        Command3.Enabled = True
        Command3.BackColor = &H8000000F
        MnuCOMM.Enabled = True
        MNUPing.Enabled = True
    Label3 = "  " & TMR
Call GoSet
Form1.Caption = CapT
Call Calculate
End Sub
'********************************************
Public Sub VOIDd() ' Empty the DATA Array
    For K = 0 To 1800000
        DATA(K) = 0
    Next K
End Sub
'********************************************
Private Sub Command8_Click() ' Zero Adjust Up
    If Zero < 11 Then Exit Sub
    Zero = Zero - 10
    Picture1.Refresh
    Call Graph
    If RecStp > 0 Then
        Call PosAdju
    End If
    Call ReDrCur
    UpDown1.Value = Zero
End Sub
'********************************************
Private Sub Command9_Click() ' Re-Zero Reference line
    Zero = 3000
    Call Graph
    If RecStp > 0 Then
        Call PosAdju
    End If
    If RUUN = False Then
        Call ReDrCur
    End If
    UpDown1.Value = Zero
End Sub
'********************************************
Private Sub Form_GotFocus()
If Form4.Visible = True Then
    Form4.Show
    Form4.SetFocus
End If
End Sub
'********************************************
Private Sub Form_Unload(Cancel As Integer) ' Exit Program
Dim NV As Integer
If RecStp > 1 And SavD = False Then
    NV = MsgBox("Do you want to save the current recording? ", vbYesNoCancel, "Save Sample?")
    If NV = vbYes Then
        Call mnuSave_Click
        NV = MsgBox("Confirm Exit Program ?", vbYesNo, "Exit?")
        If NV = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If NV = vbCancel Then
        Cancel = 1
        Exit Sub
    End If
End If
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    If Form2.Visible = True Then Unload Form2
    If Form3.Visible = True Then Unload Form3
    If Form4.Visible = True Then Unload Form4
    If Form5.Visible = True Then Unload Form5
    If Form6.Visible = True Then Unload Form6
    If Form7.Visible = True Then Unload Form7
    If Form8.Visible = True Then Unload Form8
    If Form9.Visible = True Then Unload Form9
    Set Form1 = Nothing
    Set Form2 = Nothing
    Set Form3 = Nothing
    Set Form4 = Nothing
    Set Form5 = Nothing
    Set Form6 = Nothing
    Set Form7 = Nothing
    Set Form8 = Nothing
    Set Form9 = Nothing
    Unload Me
    End
End Sub
'********************************************
Private Sub Mag1_change() ' Vertical Magnification Adjust
    MAG = (Mag1 - 50)
    Label4 = MAG
    Call PosAdju
    Form1.ReDrCur
End Sub
'********************************************
Private Sub Form_Load() ' Main Form Load
    Label14 = Format(Time, "Medium time")
    Label15 = Format(Date, "short Date")
    Timer1.Enabled = False
    Command3.Enabled = False
    PAUSE.Enabled = False
    Command7.Enabled = False
    ZAdjV = 0
    Slider1 = 0
    Label12 = ZAdjV
    Marker = 42
    YTER = 1
    Call Main
    RECORD = False
    MAG = 30
    Zero = 3000
    StepRate = 10
    ONCE = False
    Mag1 = 80
    Yr = Zero
    Yni = Zero
    Yn = Zero
    Xr = 0
    Xni = 0
    Call RENEW
    RUUN = False
    Call GoSet
    Call RENEW1
    Call TAdj
    TMR = 0
    Label3 = "  " & TMR
    Label4 = MAG
    SB1.Max = 0
    Check1 = 1
    Cursors = True
    LabAmpL = "0.0"
    LabAmpR = "0.0"
    LabDIFF = "0.00"
    LabFreq = "0.00"
    Form1.Caption = CapT & " "
    PB1.Max = 1800000
    Call RSETa
    TrigX = False
    Call RSETa
    LOrPos = CBL
    ROrPos = CBR
End Sub
'********************************************
Private Sub AdVance() 'Advance Line (Plot Graph)
Picture1.DrawWidth = DrWdt
NB = Picture1.ScaleWidth
DATA(RecStp) = TEMP
If TrigX Then
    If TEMP = TrLev Or TEMP > TrLev Then
        If Beeper = False Then
            If BpOnly Then
                Beep
            Else
                CLIK = sndPlaySound(SoundB, 1)
            End If
            Beeper = True
        End If
    Else
        Beeper = False
    End If
End If
RecStp = RecStp + 1
    Xn = Xr + StepRate
        If ((Zero - 1) - (((TEMP * MAG) - (ZAdjV * MAG)))) < 1 Then
            Yn = 1
        Else
            Yn = Zero - ((TEMP * MAG) - (ZAdjV * MAG))
        End If
    Picture1.Line (Xr, Yr)-(Xn, Yn), RGB(Gr, Gg, Gb)
    Yr = Yn
    Xr = Xn
    If Mirror Then
        Xn = Xni + StepRate
            If (Zero + ((TEMP * MAG) - (ZAdjV * MAG))) < 1 Then
            Yn = 1
        Else
            Yn = Zero + ((TEMP * MAG) - (ZAdjV * MAG))
        End If
    Picture1.Line ((Xni - 1), Yni)-((Xn - 1), Yn), RGB(Gr, Gg, Gb)
    Yni = Yn
    Xni = Xn
    End If
        If Cursors Then
            If CBL = Xn Then
                LabAmpL = TEMP
            End If
            If CBR = Xn Then
                LabAmpR = TEMP
            End If
        End If
    If Xr > (NB - 1) Then
        If ONCE = True Then
            Timer1.Enabled = False
            Call PSIM
            ONCE = False
        Else
            Xr = 0
            Xn = 0
                If RECORD Then
                    Call LOGG
                End If
            Call RENEW1
        End If
    End If
    PB1 = RecStp
Picture1.DrawWidth = 1
If RecStp > 1799998 Then Call DatEnd '1799998 - 8.7 Meg File Acquired
End Sub
'********************************************
Public Sub DatEnd()
    Call PSIM
    MsgBox "Maximum Data-Log achieved - Logging Stopped.  ", vbExclamation, "0x34's Data Logger"
End Sub
'********************************************
Public Sub LOGG()
    Frame = Frame + 1
    SB1.Max = Frame
    SB1 = SB1.Max
End Sub
'********************************************
Private Sub MnuABOUT_Click()
    Form4.Show
End Sub
'********************************************
Private Sub MNUclock_Click()
Form7.Show
End Sub
'********************************************
Private Sub MnuColor_Click()
    Form2.Show
End Sub
'********************************************
Private Sub MnuCOMM_Click()
    Form3.Show
End Sub
'********************************************
Private Sub MnuEXIT_Click() ' EXIT selected from Menu
Unload Form1
End Sub
'********************************************
Private Sub MnuGraph_Click() ' Line width adjust from menu
    Form6.Show
End Sub
'********************************************
Private Sub MNULoad_Click() ' Load File
On Error GoTo Erex
Dim TMP As Long
Dim m As String
Dim Ttemp As String
Dim Dtemp As String
Fqq = False
If SavD = False Then Call SnvNag
If Fqq Then Exit Sub
    CD1.Flags = &H1
    CD1.Flags = &H2
    CD1.DefaultExt = "dlg"
    CD1.Filter = "Logger Files (*.dlg)|*.dlg|All Files (*.*)|*.*"
    CD1.DialogTitle = "Open DataLogger File"
    CD1.Flags = &H4
    CD1.Flags = &H1000
    CD1.ShowOpen
    m = Right$(CD1.FileName, 3)
        If m = "dlg" Then
            RecStp = 0
            MousePointer = vbHourglass
            Timer3.Enabled = False
            Open CD1.FileName For Input As #1
            Input #1, Dtemp
            Input #1, Ttemp
            Input #1, TMP
                If TMP = 0 Then
                    ClockS = False
                    Form1.Check1 = 0
                    Form1.Check1.Enabled = False
                    Form1.Label1 = "Sends"
                    Form1.Label5.Visible = False
                    Form1.Label7.ForeColor = vbBlack
                    Form7.Option2 = True
                Else
                    ClockS = True
                    Form1.Check1 = 1
                    Form1.Check1.Enabled = True
                    Form1.Label1 = "Elapsed"
                    Form1.Label5.Visible = True
                    Form1.Label7.ForeColor = vbGreen
                    Form7.Option1 = True
                End If
            TMP = 0
            Do Until EOF(1)
                Input #1, TMP
                    DATA(RecStp) = TMP
                    RecStp = RecStp + 1
            Loop
            Close #1
            Label14 = Ttemp
            Label15 = Dtemp
            SavD = True
            Call TAdj
            Call PosAdju
            Call ReDrCur
            MousePointer = vbDefault
            PB1.Max = 1800000
            PB1 = RecStp
            Form1.Caption = CapT & " : " & CD1.FileName
        Else
            MsgBox "Invalid File Type Dipshit!!  ", vbCritical, "File Type Error"
        End If
    Label3 = (RecStp * 0.01)
    Exit Sub
Erex:
    MsgBox "File Load Error  ", vbInformation, "0x34 Goofed!"
    MousePointer = vbDefault
    RSETa
    Exit Sub
End Sub
'********************************************
Private Sub MNUPing_Click() ' Ping Test selected from menu
Form5.Visible = True
End Sub
'********************************************
Private Sub MnuReset_Click() ' Menu Selected RESET
Timer3.Enabled = True
Call RSETa
End Sub
'********************************************
Private Sub mnuSave_Click() ' Save Sample Recording
Dim ClStat As Integer
    On Error GoTo Errore
        If RecStp = 0 Then
            MsgBox "No Data Recorded Dipshit!!   ", vbExclamation, "Forget something?"
            Exit Sub
        End If
    PB1 = 0
    PB1.Max = RecStp
    CD1.Flags = &H1
    CD1.Flags = &H2
    CD1.DefaultExt = "dlg"
    CD1.Filter = "All Files (*.*)|*.*|Logger Files (*.dlg)|*.dlg"
    CD1.Flags = cdlOFNOverwritePrompt
    CD1.DialogTitle = "Save As DataLogger File"
    CD1.ShowSave
    MousePointer = vbHourglass
        If ClockS Then
            ClStat = 1
        Else
            ClStat = 0
        End If
    Open CD1.FileName For Output As #1
    Print #1, RecDate
    Print #1, RecTime
    Print #1, ClStat
        For UHG = 0 To RecStp
            Print #1, DATA(UHG)
            PB1 = UHG
        Next UHG
    Close #1
    SavD = True
    Form1.Caption = CapT & " : " & CD1.FileName
    MousePointer = vbDefault
    PB1.Max = 1800000
    PB1 = RecStp
Errore:
    PB1.Max = 1800000
    PB1 = RecStp
End Sub
'********************************************
Private Sub MnuSN_Click() ' Save Nag Option Select
If MnuSN.Checked Then
    MnuSN.Checked = False
    SavD = True
Else
    MnuSN.Checked = True
    SavD = False
End If
End Sub
'********************************************
Private Sub MnuTrigr_Click()
Form8.Show
End Sub
'********************************************
Private Sub MSComm1_OnComm() ' RS232 Port Monitor
TEMP = 0
FUK = ""
ModStr = ""
ModStr = MSComm1.Input
If ModStr <> "" Then
    If MSComm1.InputMode = comInputModeText Then
        FUK = Asc(ModStr)
        TEMP = TEMP + FUK
    Else
        TEMP = ModStr
    End If
Else
    FUK = "0"
    TEMP = 0
End If
If ClockS = False Then
    Call AdVance
    TMR = TMR + 1
    Label3 = "  " & TMR
End If
End Sub
'********************************************
Private Sub PAUSE_Click() ' Start / Stop
If Form8.Visible Then Unload Form8
TMR = 0
If RUUN = True Then
    RUUN = False
    Command3.Enabled = True
    Command3.BackColor = &H8000000F
    ONCE = False
    Call Comm
    MnuCOMM.Enabled = True
    MNUPing.Enabled = True
    Slider2.Enabled = True
    If Cursors Then
        CBL.Enabled = True
        CBR.Enabled = True
    End If
Else
    If RecStp > 0 Then
        Call SnvNag
    End If
    RecTime = Format(Time, "Medium time")
    RecDate = Format(Date, "short Date")
    If Joker Then Exit Sub
    Call Comm
    If MSComm1.PortOpen = False Then Exit Sub
    Call VOIDd
    RECORD = True
    Frame = 0
    SB1.Max = Frame
    RUUN = True
    RecStp = 0
    Call RENEW
    Command3.Enabled = False
    ONCE = False
    MnuCOMM.Enabled = False
    MNUPing.Enabled = False
    Slider2.Enabled = False
    CBL.Enabled = False
    CBR.Enabled = False
End If
Call GoSet
If RUUN = False Then
    Call ReDrCur
End If
End Sub
'********************************************
Public Sub SnvNag() ' Save Nag
Dim NV As Integer
Joker = False
If RecStp > 1 And SavD = False Then
    NV = MsgBox("Do you want to save the current recording? ", vbYesNoCancel, "Save Sample?")
    If NV = vbYes Then
        Call mnuSave_Click
        NV = MsgBox("Confirm Start new Recording? ", vbYesNo, "New?")
        If NV = vbNo Then
            Joker = True
            Exit Sub
        End If
    End If
    If NV = vbCancel Then
        Joker = True
        Fqq = True
        Exit Sub
    End If
End If
    If MnuSN.Checked = True Then
        SavD = False
    Else
        SavD = True
    End If
End Sub
'********************************************
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ExSbER
If Check1 And Cursors And RUUN = False Then
    If Button = 1 Then
        If X > (CBR - 1) Then Exit Sub
        If X < 1 Then Exit Sub
        CBL = X
        Posi1 = (DataChunck + X)
        LCLoc = ((Posi1 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
    If Button = 2 Then
        If X < (CBL + 1) Then Exit Sub
        If X > Picture1.ScaleWidth Then Exit Sub
        CBR = X
        Posi2 = (DataChunck + X)
        RCLoc = ((Posi2 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
End If
If Zoom Then
    LabAmpL = DATA(ZSP + CBL)
    LabAmpR = DATA(ZSP + CBR)
End If
ExSbER:
    Exit Sub
End Sub
'********************************************
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ExSbER
If Check1 And Cursors And RUUN = False Then
    If Button = 1 Then
        If X > (CBR - 1) Then Exit Sub
        If X < 1 Then Exit Sub
        CBL = X
        Posi1 = (DataChunck + X)
        LCLoc = ((Posi1 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
    If Button = 2 Then
        If X < (CBL + 1) Then Exit Sub
        If X > Picture1.ScaleWidth Then Exit Sub
        CBR = X
        Posi2 = (DataChunck + X)
        RCLoc = ((Posi2 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    End If
End If
If Zoom Then
    LabAmpL = DATA(ZSP + CBL)
    LabAmpR = DATA(ZSP + CBR)
End If
ExSbER:
    Exit Sub
End Sub
'********************************************
Private Sub SB1_Change() 'Display Posistion Adjust
    Call PosAdju
    Call ReDrCur
End Sub
'********************************************
Private Sub SB1_scroll() 'Display Posistion Adjust
    Call PosAdju
End Sub
'********************************************
Public Sub PosAdju()    ' Redraw Graph
On Error GoTo ErrorReDr
Dim A As Long
Dim NB As Long
If RUUN = True Then Exit Sub
Picture1.DrawWidth = DrWdt
    If SB1 = 0 Then
        DataChunck = 0
    Else
        DataChunck = (SB1 * Picture1.ScaleWidth)
        If (DataChuck + Picture1.ScaleWidth) > 1800000 Then
            DataChunck = DataChunck - (Picture1.ScaleWidth)
        End If
    End If
    If Zoom Then
        DataChunck = ZSP
    End If
Xr = 0
Xni = 0
    If DataChunck > 0 Then
        Yr = (Zero - (DATA(DataChunck - 1)))
        Yni = (Zero - (DATA(DataChunck - 1)))
    Else
        Yr = Zero
        Yni = Zero
    End If
Yn = Yr
Picture1.Cls
    If Check1 = 1 Then Call Graph
    If RecStp < Picture1.ScaleWidth Then
        NB = RecStp
    Else
        NB = Picture1.ScaleWidth
    End If
Picture1.DrawWidth = DrWdt
For Yx = 0 To NB
    TEMP = DATA(DataChunck)
    Xn = Xr + StepRate
        If ((Zero - 1) - (((TEMP * MAG) - (ZAdjV * MAG)))) < 1 Then
            Yn = 1
        Else
            Yn = Zero - ((TEMP * MAG) - (ZAdjV * MAG))
        End If
    Picture1.Line (Xr, Yr)-(Yx, Yn), RGB(Gr, Gg, Gb)
    Yr = Yn
    Xr = Yx
    If Mirror Then
        Xn = Xni + StepRate
            If ((Zero - 1) + (((TEMP * MAG) - (ZAdjV * MAG)))) < 1 Then
            Yn = 1
        Else
            Yn = Zero + ((TEMP * MAG) - (ZAdjV * MAG))
        End If
    Picture1.Line ((Xni - 1), Yni)-((Xn - 1), Yn), RGB(Gr, Gg, Gb)
    Yni = Yn
    Xni = Xn
    End If
        If Cursors Then
            If CBL = Yx Then
                LabAmpL = TEMP
            End If
            If CBR = Yx Then
                LabAmpR = TEMP
            End If
        End If
DataChunck = DataChunck + 1
Next Yx
Picture1.DrawWidth = 1
    Exit Sub
ErrorReDr:
    MsgBox "Some kinda f-ed up error while redrawing.", vbCritical, "0x34 Must'a Goofed"
    MsgBox "Program has become unstable and will be terminated. ", vbCritical, "Sorry"
    SavD = True
    Unload Me
End Sub
'********************************************
Private Sub Slider1_Change() ' Zero Reference Adjust
    ZAdjV = Slider1
    Label12 = ZAdjV
    If RecStp = 0 Then Exit Sub
    Call PosAdju
    Call ReDrCur
End Sub
'********************************************
Private Sub Slider1_Scroll() ' Zero Reference Adjust
    ZAdjV = Slider1
    Label12 = ZAdjV
    If RecStp = 0 Then Exit Sub
    Call PosAdju
    Call ReDrCur
End Sub
'********************************************
Private Sub Slider2_Scroll() ' Horizontal Sweep Adjust
MousePointer = vbHourglass
Call TAdj
MousePointer = vbDefault
Call ReDrCur
End Sub
'********************************************
Public Sub TZAdj() ' Adjust for ZOOM
Dim BVD As Long
Dim ScWV As Long
If Check1 Then
    CBL.Enabled = True
    CBR.Enabled = True
    LabAmpL = "0.0"
    LabAmpR = "0.0"
    LabDIFF = "0.00"
    LabFreq = "0.00"
    LabAmpL.Enabled = True
    LabAmpR.Enabled = True
    LabFreq.Enabled = True
    LabDIFF.Enabled = True
    Cursors = True
    Call Calculate
End If
    Timer1.Interval = 1
    Picture1.ScaleWidth = (Posi2 - Posi1)
    Label7 = "ZOOM"
    YTER = 0.05
If Check1 = 0 Then Cursors = False
    If Cursors Then
        CBL.Max = Picture1.ScaleWidth
        CBR.Max = Picture1.ScaleWidth
        CBL = 1
        CBR = (Picture1.ScaleWidth - 1)
    Else
        CBL.Enabled = False
        CBR.Enabled = False
        LabAmpL = "0.0"
        LabAmpR = "0.0"
        LabDIFF = "0.00"
        LabFreq = "0.00"
        LabAmpL.Enabled = False
        LabAmpR.Enabled = False
        LabFreq.Enabled = False
        LabDIFF.Enabled = False
        Cursors = False
    End If
Call Graph
Call PosAdju
End Sub
'********************************************
Public Sub TAdj() ' Horiz Select and Adjust
Dim BVD As Long
If Check1 Then
    CBL.Enabled = True
    CBR.Enabled = True
    LabAmpL = "0.0"
    LabAmpR = "0.0"
    LabDIFF = "0.00"
    LabFreq = "0.00"
    LabAmpL.Enabled = True
    LabAmpR.Enabled = True
    LabFreq.Enabled = True
    LabDIFF.Enabled = True
    Cursors = True
    Call Calculate
End If
If Check1 And Zoom = False Then
    Command5.Enabled = True
End If
If Slider2 = 0 Then
        If RecStp > 1750000 Then
            MsgBox "10mS review is not possible with a recording of this size.  ", vbInformation, "0x34"
            Slider2 = 2
            Call TAdj
            Exit Sub
        End If
    Timer1.Interval = 1
    Picture1.ScaleWidth = 10
    StepRate = 1
    Label7 = "10 mS "
    YTER = 0.01
    Cursors = True
    Command5.Enabled = False
End If
If Slider2 = 1 Then
        If RecStp > 1750000 Then
            MsgBox "50mS review is not possible with a recording of this size.  ", vbInformation, "0x34"
            Slider2 = 2
            Call TAdj
            Exit Sub
        End If
    Timer1.Interval = 1
    Picture1.ScaleWidth = 50
    StepRate = 1
    Label7 = "50 mS "
    YTER = 0.05
    Cursors = True
    Command5.Enabled = False
End If
If Slider2 = 2 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 100
    StepRate = 1
    Label7 = "100 mS "
    YTER = 0.1
    Cursors = True
End If
If Slider2 = 3 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 500
    StepRate = 1
    Label7 = "500 mS "
    YTER = 0.5
    Cursors = True
End If
If Slider2 = 4 Then
    VFq = 3.5
    Timer1.Interval = 1
    Picture1.ScaleWidth = 1000
    StepRate = 1
    Label7 = "1 Sec "
    YTER = 1
    Cursors = True
End If
If Slider2 = 5 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 5000
    StepRate = 1
    Label7 = "5 Sec "
    YTER = 5
    Cursors = True
End If
If Slider2 = 6 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 10000
    StepRate = 1
    Label7 = "10 Sec "
    YTER = 10
    Cursors = True
End If
If Slider2 = 7 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 30000
    StepRate = 1
    Label7 = "30 Sec "
    Cursors = False
    Command5.Enabled = False
End If
If Slider2 = 8 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 60000
    StepRate = 1
    Label7 = "1 Min "
    Cursors = False
    Command5.Enabled = False
End If
If Slider2 = 9 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 600000
    StepRate = 1
    Label7 = "10 Min "
    Cursors = False
    Command5.Enabled = False
End If
If Slider2 = 10 Then
    Timer1.Interval = 1
    Picture1.ScaleWidth = 1800000
    StepRate = 1
    Label7 = "30 Min "
    Cursors = False
    Command5.Enabled = False
End If
If Check1 = 0 Then Cursors = False
    If Cursors Then
        CBL.Max = Picture1.ScaleWidth
        CBR.Max = Picture1.ScaleWidth
        If Zoom = False Then
            CBL = 1
            CBR = (Picture1.ScaleWidth - 1)
        End If
        Posi1 = (DataChunck + CBL)
        LCLoc = ((Posi1 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
        Posi2 = (DataChunck + CBR)
        RCLoc = ((Posi2 - (DataChunck - 1)) + (Picture1.ScaleWidth * SB1))
    Else
        CBL.Enabled = False
        CBR.Enabled = False
        LabAmpL = "0.0"
        LabAmpR = "0.0"
        LabDIFF = "0.00"
        LabFreq = "0.00"
        LabAmpL.Enabled = False
        LabAmpR.Enabled = False
        LabFreq.Enabled = False
        LabDIFF.Enabled = False
        Cursors = False
    End If
Call Graph
If RecStp > 0 Then
    BVD = Round((RecStp / Picture1.ScaleWidth), 0)
    If BVD < 1 Then BVD = 0
    SB1.Max = BVD
    Call PosAdju
End If
End Sub
'********************************************
Private Sub Timer1_Timer() ' Main Timer
If ClockS Then
    Call AdVance
    TMR = TMR + 0.01
    Label3 = "  " & TMR
End If
End Sub
'********************************************
Private Sub RENEW()
    Dim NB As Long
    NB = Picture1.ScaleWidth
    Picture1.Refresh
    Call Graph
    Yr = Zero
    Yn = Zero
    Xr = 0
Call ReDrCur
End Sub
'********************************************
Private Sub RENEW1() ' Clear Data Array
    Dim NB As Long
    NB = Picture1.ScaleWidth
    Picture1.Refresh
    Call Graph
    Xr = 0
End Sub
'********************************************
Private Sub GoSet() ' Initialize
If RUUN = True Then
    PAUSE.BackColor = vbGreen
    Timer1.Enabled = True
Else
    PAUSE.BackColor = vbRed
    Timer1.Enabled = False
End If
End Sub
'********************************************
Public Sub Graph() ' Draw Grid, Zero Reference Line, and Markers
Dim Mrk As Long
Dim POSI As Long
Dim CntLine As Long
If Check1 = 1 Then
    Picture1.Refresh
    Picture1.Cls
    Picture1.DrawStyle = 2
    Dim NB As Long
    Dim Va As Long
    Dim Vb As Long
    Dim CENT As Long
    Va = 0
    Vb = 0
    Picture1.DrawWidth = 1
    NB = Picture1.ScaleWidth
    Va = (Picture1.ScaleWidth / 10)
    Picture1.Line (0, 1800)-(NB, 1800), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 2400)-(NB, 2400), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 1200)-(NB, 1200), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 600)-(NB, 600), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 3600)-(NB, 3600), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 4200)-(NB, 4200), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 4800)-(NB, 4800), RGB(Cr, Cg, Cb)
    Picture1.Line (0, 5400)-(NB, 5400), RGB(Cr, Cg, Cb)
    Picture1.Line (Va, 0)-(Va, 6000), RGB(Cr, Cg, Cb)
        Va = Va + Vb
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
      Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.DrawStyle = 0
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB((Cr + 20), (Cg + 20), (Cb + 20))
    Picture1.DrawStyle = 2
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
        Vb = Vb + Va
    Picture1.Line (Vb, 0)-(Vb, 6000), RGB(Cr, Cg, Cb)
    Picture1.DrawStyle = 0
    Picture1.Line (0, 3000)-(NB, 3000), RGB(Cr, Cg, Cb)
'                               Draw ZERO BAR
        Picture1.Line (0, Zero)-(NB, Zero), RGB(CZr, CZg, CZb)
    If (Marker - 12) > 0 Then ' Zero Line Markers
        Mrk = (Picture1.ScaleWidth / 50)
        POSI = 0
            For KHG = 0 To 50
                Picture1.Line (POSI, (Zero + (Marker - 9)))-(POSI, (Zero - Marker)), RGB(CZr, CZg, CZb)
                POSI = POSI + Mrk
            Next KHG
    End If
Else
    Picture1.Refresh
    Picture1.Cls
End If
If Cursors = True Then
    Call PlaceCur
End If
End Sub
'********************************************
Public Sub PlaceCur() ' Draw Cursors
On Error GoTo JIGGY
    Picture1.DrawMode = 7
    Picture1.DrawStyle = 2
    Picture1.Line (CR1a, 0)-(CR1a, 6000), RGB(CLRx, CLGx, CLBx)
    Picture1.Line (CR2a, 0)-(CR2a, 6000), RGB(CRRx, CRGx, CRBx)
    CR1a = CBL
    CR2a = CBR
    Picture1.Line (CBL, 0)-(CBL, 6000), RGB(CLRx, CLGx, CLBx)
    Picture1.Line (CBR, 0)-(CBR, 6000), RGB(CRRx, CRGx, CRBx)
    Picture1.DrawMode = 13
    If Zoom = False Then
        Posi1 = (DataChunck + CBL)
        LCLoc = ((Posi1 - (DataChunck)) + (Picture1.ScaleWidth * SB1))
        Posi2 = (DataChunck + CBR)
        RCLoc = ((Posi2 - (DataChunck)) + (Picture1.ScaleWidth * SB1))
        LabAmpL = DATA(LCLoc)
        LabAmpR = DATA(RCLoc)
    End If
    Picture1.DrawStyle = 0
    Exit Sub
JIGGY: 'I'm Jiggy wit it!
    Exit Sub
End Sub
'********************************************
Public Sub ReDrCur() ' ReDraw the Cursors when moved
Dim Weight As Integer
Weight = Picture1.DrawWidth
If RUUN Then Exit Sub
If Cursors = False Then Exit Sub
Picture1.DrawWidth = 1
Picture1.DrawStyle = 2
    If Cursors = True Then
        Picture1.DrawMode = 7
        Picture1.Line (CBL, 0)-(CBL, 6000), RGB(CLRx, CLGx, CLBx)
        Picture1.Line (CBR, 0)-(CBR, 6000), RGB(CRRx, CRGx, CRBx)
        Picture1.DrawMode = 13
    End If
Picture1.DrawStyle = 0
Picture1.DrawWidth = Weight
End Sub
'********************************************
Private Sub Timer2_Timer() ' Splash Screen Control and GRID start-up
Slider2.Enabled = False
    If MBV = False Then
        Call Graph
        MBV = True
        Form4.Visible = True
        Form4.SetFocus
        Command3.Enabled = False
        Command2.Enabled = False
    Else
        ASH = ASH + 1
            If ASH > 20 Then
                Form4.Visible = False
                Unload Form4
                Timer2.Enabled = False
            End If
    End If
If Timer2.Enabled = False Then
    CBR = (Picture1.ScaleWidth - 20)
    CBL = 20
    Slider2.Enabled = True
    Call RSETa
    Command3.Enabled = True
    PAUSE.Enabled = True
    Command7.Enabled = True
    Command2.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
    Label14 = Format(Time, "Medium time")
End Sub

'********************************************
Private Sub UpDown1_Change() ' Zero Line Adjust
    Zero = UpDown1
    Call Graph
    If RecStp > 0 Then
        Call PosAdju
    End If
    If RUUN = False Then
        Call ReDrCur
    End If
End Sub

'                             Code by 0x34 - 2004
