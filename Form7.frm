VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Clock Source"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Use External Clock Source"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   460
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Internal Clock Source"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   210
      Width           =   2295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim HKK As String
If Option2 Then
    If ClockS = True Then
        If RecStp > 1 And SavD = False Then
            HKK = "                 **  WARNING  **" & vbNewLine
            HKK = HKK & "ALL RECORDED DATA WILL BE LOST!!" & vbNewLine & vbNewLine
            HKK = HKK & "                      Continue?    "
            Don = MsgBox(HKK, vbYesNo, "Switch Clock")
            If Don <> 6 Then
                GoTo Bennt
            End If
        End If
        ClockS = False
        Form1.Check1 = 0
        Form1.Command5.Enabled = False
        Form1.Check1.Enabled = False
        Form1.Label1 = "Sends"
        Form1.Label5.Visible = False
        Form1.Label7.ForeColor = vbBlack
        SavD = True
        Call Form1.RSETa
    End If
Else
    If ClockS = False Then
        If RecStp > 1 And SavD = False Then
            HKK = "                 **  WARNING  **" & vbNewLine
            HKK = HKK & "ALL RECORDED DATA WILL BE LOST!!" & vbNewLine & vbNewLine
            HKK = HKK & "                      Continue?    "
            Don = MsgBox(HKK, vbYesNo, "Switch Clock")
            If Don <> 6 Then
                GoTo Bennt
            End If
        End If
        ClockS = True
        Form1.Check1 = 1
        Form1.Command5.Enabled = True
        Form1.Check1.Enabled = True
        Form1.Label1 = "Elapsed"
        Form1.Label5.Visible = True
        Form1.Label7.ForeColor = vbGreen
        SavD = True
        Call Form1.RSETa
    End If
End If
Unload Me
Bennt:
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If ClockS = True Then
    Option1 = True
Else
    Option2 = True
End If
End Sub

'                             Code by 0x34 - 2004
