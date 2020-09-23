VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VisualEQ 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3645
   ClientLeft      =   1560
   ClientTop       =   3060
   ClientWidth     =   4920
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmVisEQ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVisEQ.frx":0ABA
   ScaleHeight     =   3645
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Text            =   "Draw width"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   6
      Left            =   240
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1200
   End
   Begin MSComctlLib.ProgressBar ProgressBar10 
      Height          =   1800
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "16kHz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar9 
      Height          =   1800
      Left            =   3390
      TabIndex        =   1
      ToolTipText     =   "12kHz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar8 
      Height          =   1800
      Left            =   3090
      TabIndex        =   2
      ToolTipText     =   "6kHz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar7 
      Height          =   1800
      Left            =   2805
      TabIndex        =   3
      ToolTipText     =   "3kHz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar6 
      Height          =   1800
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "1kHz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar5 
      Height          =   1800
      Left            =   2235
      TabIndex        =   5
      ToolTipText     =   "31Hz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar4 
      Height          =   1800
      Left            =   1950
      TabIndex        =   6
      ToolTipText     =   "62Hz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   1800
      Left            =   1650
      TabIndex        =   7
      ToolTipText     =   "125Hz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   1800
      Left            =   1365
      TabIndex        =   8
      ToolTipText     =   "250Hz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   1800
      Left            =   1080
      TabIndex        =   9
      ToolTipText     =   "500Hz"
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Play Something In any program.. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   0
      MousePointer    =   3  'I-Beam
      Picture         =   "frmVisEQ.frx":0EFC
      Top             =   0
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   240
      X2              =   4080
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "VisualEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sa, sb, sc, sd, se, sf, sg, Sh, si, sj

Dim hmixer As Long                      ' mixer handle
Dim inputVolCtrl As MIXERCONTROL        ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL       ' microphone volume control
Dim rc As Long                          ' return code
Dim OK As Boolean                       ' boolean return code
Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' Volume Buffer
Private VU As VULights                  ' Volume Unit Values
Private FreqNum As Frequency
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub VolVal(VolIs As Long, VolFreq As Double)
For FreqNum = 0 To 9
Next FreqNum
VolIs = volume * 327.67
VolFreq = VU.Freq(FreqNum)
VU.FreqVal = VolIs * VolFreq
End Sub

Private Sub LightsA()
Dim a

' ProgressBar1
FreqNum = Freq500Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.2) To FreqNum
Next VU.InOutLev
a = VU.InOutLev - 500
sa = a * 125

ProgressBar1.Value = VU.InOutLev
End Sub

Private Sub LightsB()
Dim b

' ProgressBar2
FreqNum = Freq250Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.4) To FreqNum
Next VU.InOutLev
b = VU.InOutLev - 250
sb = b * 125

ProgressBar2.Value = VU.InOutLev
End Sub

Private Sub LightsC()
Dim c

' ProgressBar3
FreqNum = Freq125Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.8) To FreqNum
Next VU.InOutLev
c = VU.InOutLev - 125
sc = c * 125

ProgressBar3.Value = VU.InOutLev
End Sub

Private Sub LightsD()
Dim d

' ProgressBar4
FreqNum = Freq62Hz
For VU.InOutLev = CDbl(VU.VolLev * 1.61290322580645E-02) To FreqNum
Next VU.InOutLev
d = VU.InOutLev - 62.5
sd = d * 125
ProgressBar4.Value = VU.InOutLev
End Sub
Private Sub LightsE()
Dim e

' ProgressBar5
FreqNum = Freq31Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.032258064516129) To FreqNum
Next VU.InOutLev
e = VU.InOutLev - 31.2
se = e * 125
ProgressBar5.Value = VU.InOutLev
End Sub

Private Sub LightsF()
Dim f
' ProgressBar6
FreqNum = Freq1kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.01) To FreqNum
Next VU.InOutLev
f = VU.InOutLev - 1
sf = f * 125
ProgressBar6.Value = VU.InOutLev
End Sub

Private Sub LightsG()
Dim g

' ProgressBar7
FreqNum = Freq3kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.03) To FreqNum
Next VU.InOutLev
g = VU.InOutLev - 3
sg = g * 125
ProgressBar7.Value = VU.InOutLev
End Sub

Private Sub LightsH()
Dim h

' ProgressBar8
FreqNum = Freq6kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.06) To FreqNum
Next VU.InOutLev
h = VU.InOutLev - 6
Sh = h * 125
ProgressBar8.Value = VU.InOutLev
End Sub

Private Sub LightsI()
Dim i

' ProgressBar9
FreqNum = Freq12kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.12) To FreqNum
Next VU.InOutLev
i = VU.InOutLev - 12
si = i * 125

ProgressBar9.Value = VU.InOutLev
End Sub

Private Sub LightsJ()
Dim j

' ProgressBar10
FreqNum = Freq16kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.16) To FreqNum
Next VU.InOutLev
j = VU.InOutLev - 16
sj = j * 125

ProgressBar10.Value = VU.InOutLev
End Sub


 

Private Sub Combo1_Click()
On Error Resume Next
Me.DrawWidth = Combo1.Text

End Sub

Private Sub Form_Load()
On Error Resume Next

MsgBox "This is based upon a code I found on PSC (forgive me for not knowing the name of the author), that showed a visual equalizer in progress bars - I haven't changed that, but I added a visualisation of the sound using RGB colors to express different frequences -- this can be used as a base to create more advanced visualisations for sound", vbInformation, "Michael Belenky"

Combo1.AddItem "1"
 Combo1.AddItem "2"
  Combo1.AddItem "3"
   Combo1.AddItem "4"





Timer2.Interval = 6

  '  PictureClip1.Cols = 2
  '  PictureClip1.Rows = 1
  '  CloseButton.Picture = PictureClip1.GraphicCell(0)
  '  AboutButton.Picture = PictureClip2.GraphicCell(2)
    Timer1.Interval = 6
   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   ' Get the output volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (OK = True) Then
   ' Set frequencies for the Volume Units
    ProgressBar1.Max = Frequency.Freq500Hz + 1
    ProgressBar1.Min = Frequency.Freq500Hz
    ProgressBar2.Max = Frequency.Freq250Hz + 1
    ProgressBar2.Min = Frequency.Freq250Hz
    ProgressBar3.Max = Frequency.Freq125Hz + 1
    ProgressBar3.Min = Frequency.Freq125Hz
    ProgressBar4.Max = Frequency.Freq62Hz + 1
    ProgressBar4.Min = Frequency.Freq62Hz
    ProgressBar5.Max = Frequency.Freq31Hz + 1
    ProgressBar5.Min = Frequency.Freq31Hz
    ProgressBar6.Max = Frequency.Freq1kHz + 1
    ProgressBar6.Min = Frequency.Freq1kHz
    ProgressBar7.Max = Frequency.Freq3kHz + 1
    ProgressBar7.Min = Frequency.Freq3kHz
    ProgressBar8.Max = Frequency.Freq6kHz + 1
    ProgressBar8.Min = Frequency.Freq6kHz
    ProgressBar9.Max = Frequency.Freq12kHz + 1
    ProgressBar9.Min = Frequency.Freq12kHz
    ProgressBar10.Max = Frequency.Freq16kHz + 1
    ProgressBar10.Min = Frequency.Freq16kHz
   Else
      MsgBox "Couldn't get waveout meter"
   End If
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Image1_Click()
MsgBox "Visualisation by Michael Belenky, freq detection code found on PSC"
End

End Sub

Private Sub Timer1_Timer()
    
    VU.VolLev = volume / 327.67
    If (volume < 0) Then volume = -volume
    ' Get the current output level
    If (1 = 1) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    'MsgBox (mxcd.cChannels)
    
  If (volume < 0) Then volume = -volume
  End If
    ActivateVolumeUnits
End Sub

Private Sub ActivateVolumeUnits()
    LightsA
    LightsB
    LightsC
    LightsD
    LightsE
    LightsF
    LightsG
    LightsH
    LightsI
    LightsJ
End Sub



Private Sub Timer2_Timer()

Dim aboutstr, edited As String
Dim i, stl

      Circle (Me.ScaleWidth / 2, 1200), (800), RGB((sa + sb + sc) / 3, (se + sf + sg) / 4, (Sh + si + sj) / 3)
     Circle (Me.ScaleWidth / 2 + 100, 1220), (700), RGB((sa + sb + sc) / 3, (se + sf + sg) / 4, (Sh + si + sj) / 3)
    Circle (Me.ScaleWidth / 2 + 200, 1250), (600), RGB((sa + sb + sc) / 2, (se + sf + sg) / 3, (Sh + si + sj) / 2)
   Circle (Me.ScaleWidth / 2 + 220, 1270), (500), RGB((sa + sb + sc) / 2, (se + sf + sg) / 3, (Sh + si + sj) / 2)
  Circle (Me.ScaleWidth / 2 + 300, 1300), (400), RGB((sa + sb + sc), (se + sf + sg) / 2, (Sh + si + sj))
 Circle (Me.ScaleWidth / 2 + 320, 1320), (300), RGB((sa + sb + sc), (se + sf + sg), (Sh + si + sj))
Circle (Me.ScaleWidth / 2 + 330, 1340), (200), RGB((sa + sb + sc), (se + sf + sg) / 2, (Sh + si + sj))
For i = 0 To 200
'Line (Line1.X1, Line1.Y1)-(Line1.X2 + (sb), Line1.Y2 - (sa * 5)), RGB((sa + sb + sc), (sd + se + sf + sg) / 2, (Sh + si + sj))

'Line (Line1.X1, Line1.Y1)-(Line1.X2 + (sb), Line1.Y2 - (sa * 5)), vbBlack
Line (sa * 50, sb * 50)-(Sh * 50, sj * 50), RGB((sa + sb + sc), (sd + se + sf + sg) / 2, (Sh + si + sj))
Line (sa * 50, sb * 50)-(Sh * 50, sj * 50), vbBlack
Next i

End Sub
