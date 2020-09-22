Attribute VB_Name = "Module1"
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

Global Cr As Integer
Global Cg As Integer
Global Cb As Integer
Global Gr As Integer
Global Gg As Integer
Global Gb As Integer
Global CZr As Integer
Global CZg As Integer
Global CZb As Integer
Global CBR As Integer
Global CBg As Integer
Global CBb As Integer
Global CLRx As Integer
Global CLGx As Integer
Global CLBx As Integer
Global CRRx As Integer
Global ZSP As Long
Global CRGx As Integer
Global ZAdjV As Integer
Global LCLoc As Long
Global RCLoc As Long
Global CRBx As Integer
Global CONF As String
Global BRate As String
Global SavD As Boolean
Global Joker As Boolean
Global LOrPos As Long
Global ROrPos As Long
Global Parity As String
Global StopBits As String
Global NBytes As String
Global PrtNumb As Integer
Global DrWdt As Integer
Global RECORD As Boolean
Global Frame As Integer
Global Beeper As Boolean
Global ZEFin As Boolean
Global SoundB As String
Global DATA(1800000) As Long
Global RecStp As Long
Global Marker As Integer
Global BpOnly As Boolean
Global Cursors As Boolean
Global CurL As Long
Global CurR As Long
Global YTER As Long
Global ClockS As Boolean
Global RUUN As Boolean
Global CR1a As Long
Global CR2a As Long
Global Xni As Long
Global Yni As Long
Global ASq As Integer
Global TrLev As Integer
Global TrigX As Boolean
Global Fqq As Boolean
Global Zero As Integer
Global TEMP As Byte
Global ASH As Integer
Global RecTime As String
Global RecDate As String
Global CapT As String
Global Mirror As Boolean
Global Zoom As Boolean
Global DataChunck As Long
Global Posi1 As Long
Global Posi2 As Long
Global REXsav As String ' Sound Info Save File
Global CustCol As String ' Custom Color settings save file
Global CommSett As String 'Comm Save File
' The one and only API call
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" ( _
    ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const csndsync = &H0, csndasync = &H1
Const csndnodefault = &H2
Const csndloop = &H8, csndnostop = &H10

Public Sub Main()
REXsav = "C:\DLogSnd.cfg"
CustCol = "C:\DLogC.cfg"
CommSett = "C:\DLogCm.cfg"
Beeper = False
ASq = 0
CR1a = 20
CR2a = 980
CBR = 980
CBL = 20
ClockS = True
YTER = 1
CLRx = 0
CLGx = 255
CLBx = 255
CRRx = 255
CRGx = 255
CRBx = 0
Cr = 0
Cg = 50
Cb = 100
Gr = 0
Gg = 255
Gb = 0
CZr = 255
CZg = 0
CZb = 0
CBR = 0
CBg = 0
CBb = 0
DrWdt = 1
CapT = ""
CapT = CapT + Chr(&HA9 - &H65)
CapT = CapT + Chr(&H8F - &H2E)
CapT = CapT + Chr(&HBF - &H4B)
CapT = CapT + Chr(&H8B - &H2A)
CapT = CapT + Chr(&H74 - &H28)
CapT = CapT + Chr(&H9E - &H2F)
CapT = CapT + Chr(&H94 - &H2D)
CapT = CapT + Chr(&H91 - &H2A)
CapT = CapT + Chr(&H90 - &H2B)
CapT = CapT + Chr(&HAC - &H3A)
CapT = CapT + Chr(&H71 - &H51)
CapT = CapT + Chr(&H4B - &H2B)
CapT = CapT + Chr(&H4F - &H22)
CapT = CapT + Chr(&H42 - &H22)
CapT = CapT + Chr(&H3F - &H1F)
CapT = CapT + Chr(&H9E - &H5C)
CapT = CapT + Chr(&H8F - &H16)
CapT = CapT + Chr(&H72 - &H52)
CapT = CapT + Chr(&H7A - &H4A)
CapT = CapT + Chr(&H90 - &H18)
CapT = CapT + Chr(&H76 - &H43)
CapT = CapT + Chr(&H69 - &H35)
Call LCFG
On Error GoTo Yelling
    Open CommSett For Input As #1
        Input #1, BRate
        Input #1, Parity
        Input #1, NBytes
        Input #1, StopBits
        Input #1, PrtNumb
    Close #1
    CONF = BRate & "," & Parity & "," & NBytes & "," & StopBits
Exit Sub
Yelling:
    BRate = "9600"
    Parity = "N"
    StopBits = "1"
    NBytes = "8"
    PrtNumb = 1
    CONF = BRate & "," & Parity & "," & NBytes & "," & StopBits

End Sub

Public Sub LCFG()
Dim Ben As Integer
On Error GoTo TError
    Open REXsav For Input As #1
    Input #1, SoundB
    Input #1, Ben
    Input #1, TrLev
    Close #1
    If Ben = 1 Then
        BpOnly = True
    Else
        BpOnly = False
    End If
Exit Sub
TError:
    MsgBox "Configuration File not found." & vbNewLine & "A new file will be created.", vbInformation, "0x34"
    Close #1
    Open REXsav For Output As #1
    Print #1, "C:\Windows\Media\Utopia Critical Stop.wav"
    Print #1, 0
    Print #1, 35
    Close #1
SoundB = "C:\Windows\Media\Utopia Critical Stop.wav"
BpOnly = False
TrLev = 35
End Sub

Public Sub ColSav()
    Open CustCol For Output As #1
        Print #1, CLRx
        Print #1, CLGx
        Print #1, CLBx
        Print #1, CRRx
        Print #1, CRGx
        Print #1, CRBx
        Print #1, Cr
        Print #1, Cg
        Print #1, Cb
        Print #1, Gr
        Print #1, Gg
        Print #1, Gb
        Print #1, CZr
        Print #1, CZg
        Print #1, CZb
        Print #1, CBR
        Print #1, CBg
        Print #1, CBb
    Close #1
End Sub

Public Sub ColLoad()
On Error GoTo YError
    Open CustCol For Input As #1
        Input #1, CLRx
        Input #1, CLGx
        Input #1, CLBx
        Input #1, CRRx
        Input #1, CRGx
        Input #1, CRBx
        Input #1, Cr
        Input #1, Cg
        Input #1, Cb
        Input #1, Gr
        Input #1, Gg
        Input #1, Gb
        Input #1, CZr
        Input #1, CZg
        Input #1, CZb
        Input #1, CBR
        Input #1, CBg
        Input #1, CBb
    Close #1
Exit Sub
YError:
    MsgBox "No Custom Color File Located!", vbInformation, "DataLogger - 0x34"
End Sub

'                             Code by 0x34 - 2004
