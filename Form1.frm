VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Morse Trainer"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5970
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer CursorPosTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   6510
   End
   Begin VB.Frame Frame2 
      Caption         =   "Play random text"
      Height          =   945
      Left            =   90
      TabIndex        =   7
      Top             =   6000
      Width           =   5685
      Begin VB.OptionButton optLetters 
         Caption         =   "All"
         Height          =   345
         Index           =   2
         Left            =   4710
         TabIndex        =   13
         Top             =   390
         Width           =   855
      End
      Begin VB.OptionButton optLetters 
         Caption         =   "Numbers"
         Height          =   345
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   390
         Width           =   945
      End
      Begin VB.OptionButton optLetters 
         Caption         =   "Letters"
         Height          =   345
         Index           =   0
         Left            =   2730
         TabIndex        =   11
         Top             =   390
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   525
         Left            =   1410
         TabIndex        =   9
         Top             =   270
         Width           =   1125
      End
      Begin VB.CommandButton cmdPlayRandom 
         Caption         =   "Play"
         Height          =   525
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Play entered text"
      Height          =   945
      Left            =   90
      TabIndex        =   4
      Top             =   5010
      Width           =   5685
      Begin VB.CommandButton cmdStopPlay 
         Caption         =   "Stop"
         Height          =   525
         Left            =   3900
         TabIndex        =   10
         Top             =   270
         Width           =   1600
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   525
         Left            =   2010
         TabIndex        =   6
         Top             =   270
         Width           =   1600
      End
      Begin VB.CommandButton cmdEncode 
         Caption         =   "Text -> Morse"
         Height          =   525
         Left            =   150
         TabIndex        =   5
         Top             =   270
         Width           =   1600
      End
   End
   Begin MSComctlLib.Slider slSpeedAdjust 
      Height          =   525
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Set Wpm"
      Top             =   4080
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   926
      _Version        =   393216
      BorderStyle     =   1
      Min             =   1
      SelStart        =   3
      Value           =   3
   End
   Begin VB.TextBox txtMorse 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1110
      Width           =   5685
   End
   Begin VB.TextBox txtEnglish 
      Height          =   885
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0CCA
      Top             =   150
      Width           =   5685
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  4            8         12         16         20         24         28        32          36      wpm "
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4620
      Width           =   5625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEncode_Click()

txtMorse.Text = ""

If txtEnglish.Text <> "" Then
    toMorse txtEnglish.Text
End If

End Sub

Private Sub cmdPlay_Click()

If txtEnglish.Text = "" Or txtMorse.Text = "" Then Exit Sub

cmdPlay.Enabled = False

EnteredTextPlaying = True

RandomTerminated = False

Form1.txtMorse.SelStart = 0
Form1.txtMorse.SelLength = 0
Form1.txtMorse.SetFocus

PlayEnteredText

End Sub

Private Sub cmdPlayRandom_Click()

cmdPlayRandom.Enabled = False

RandomTerminated = False

RandomPlaying = True

EnteredTextPlaying = False

txtEnglish.Text = ""

Form1.txtMorse.SelStart = 0
Form1.txtMorse.SelLength = 0
Form1.txtMorse.SetFocus

PlayRandomText

End Sub


Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If RandomPlaying = False Then Exit Sub

RandomTerminated = True

DSB.Stop

RandomPlaying = False

End Sub


Private Sub cmdStopPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If EnteredTextPlaying = False Then Exit Sub

DSB.Stop

CursorPosTimer.Enabled = False

EnteredTextPlaying = False

cmdPlay.Enabled = True

End Sub

Private Sub CursorPosTimer_Timer()
Dim TotalSize As Long, StringSize As Long


TotalSize = UBound(sample)              'get the size of the soundbuffer

StringSize = Len(txtMorse.Text)         'get the size of the morse text string

DSB.GetCurrentPosition dwPlayCursor     'get the current position in the soundbuffer

txtMorse.SetFocus                       'need focus on the textbox in order to be able
                                        'to highlight text

If dwPlayCursor.lPlay > 0 Then  'once playing has started, highlight the text
    txtMorse.SelStart = 0       'as playing progresses
    txtMorse.SelLength = StringSize * (dwPlayCursor.lPlay / TotalSize)
End If

If dwPlayCursor.lPlay = 0 Then  'playing of text is completed
    CursorPosTimer.Enabled = False
    txtMorse.SelStart = 0
    txtMorse.SelLength = 0
    'txtMorse.Refresh
    PlayingCompleted = True
    cmdPlay.Enabled = True
End If

If RandomTerminated = True Then 'playing of random text is stopped
    CursorPosTimer.Enabled = False
    txtMorse.SelStart = 0
    txtMorse.SelLength = 0
    PlayingCompleted = True
    cmdPlayRandom.Enabled = True
End If

End Sub


Private Sub Form_Load()

Initialise_MorseTables              'fill conversion tables with values

SpeedFactor = slSpeedAdjust.Value   'default speed is 12 wpm

PlayType = LETTERS                  'default character generation is letters only

Set DS = DX7.DirectSoundCreate(vbNullString)    'initialize directsound
DS.SetCooperativeLevel hWnd, DSSCL_NORMAL
PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = 4410 * SpeedFactor / 3
PCM.nBitsPerSample = 8
PCM.nBlockAlign = 1
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
DSBD.lFlags = DSBCAPS_STATIC


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next    'if no sound is being played dsb.stop might raise error

DSB.Stop

RandomTerminated = True

Set DSB = Nothing

Set DS = Nothing

Unload Me

End Sub

Private Sub Form_Resize()

If Form1.WindowState = vbMinimized Then Exit Sub

Form1.Width = 6090

Form1.Height = 7500

End Sub

Private Sub optLetters_Click(Index As Integer)

Select Case Index
    Case 0: PlayType = LETTERS
    Case 1: PlayType = NUMBERS
    Case 2: PlayType = ALL
End Select

End Sub

Private Sub slSpeedAdjust_Click()

SpeedFactor = slSpeedAdjust.Value

PCM.lSamplesPerSec = 4410 * SpeedFactor / 3
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign

End Sub
