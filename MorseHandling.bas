Attribute VB_Name = "MorseHandling"
Option Explicit
Public sPos As Long
Public sample() As Byte                 'the buffer that holds the morse sound
Public timeunit As Long
Public hMorse(52, 1) As String          'conversion table text and morse
Public PlayType As Integer              'for choosing characters in random training
Public Const LETTERS = 25
Public Const NUMBERS = 35
Public Const ALL = 52
Public SpeedFactor As Long              'for adjusting the playback speed
Public RandomTerminated As Boolean      'to indicate random text play is terminated
Public PlayingCompleted As Boolean      'to indicate that entered text play is terminated
Public RandomPlaying As Boolean         'used for determining which stop button to use
Public EnteredTextPlaying As Boolean    'used for determining which stop button to use
Public dwPlayCursor As DSCURSORS        'for obtaining the current position in the soundbuffer
Public DX7 As New DirectX7
Public DS As DirectSound
Public DSB As DirectSoundBuffer
Public PCM As WAVEFORMATEX
Public DSBD As DSBUFFERDESC

Public Function makeSample(morse As String) As Byte
Dim ditLen As Integer, length As Long
Dim dahlen As Integer
Dim wordLen As Long
Dim charPause As Integer
Dim wordPause As Integer
Dim wpm As Long, extra As Integer, compression As Double
Dim c As String, i As Integer

    wpm = 25
    ditLen = 1
    dahlen = 3
    wordLen = 50
    charPause = 3
    wordPause = 7

    ' milliseconds per dit
    timeunit = 1000 * 6 / (wpm * wordLen)
    timeunit = timeunit * 10
    If (timeunit > 100) Then
        ' If the characters are to be sent slowly (less than 12wpm)
        ' compress the elements, and stretch the gap between characters
        ' to keep the wpm the same.
        compression = timeunit / 100 - 1
        timeunit = 100
    End If
    If (timeunit <= 100) Then
        compression = 0
    End If

    length = 0
    extra = 0
    For i = 1 To Len(morse)
        c = Mid(morse, i, 1)
        Select Case c
            Case ".":   extra = extra + ditLen + ditLen
                        length = length + ditLen + ditLen
            Case "-":   extra = extra + dahlen + ditLen
                        length = length + dahlen + ditLen
            Case " ":   extra = extra + charPause - ditLen
                        length = length + charPause - ditLen + (extra * compression)
                        extra = 0
            Case "/":   extra = extra + wordPause - ditLen
                        length = length + wordPause - ditLen + (extra * compression)
                        extra = 0
        End Select
    Next i
    
    length = length * 8 * timeunit

    ReDim sample(length)
    
    sPos = 0

    extra = 0
    For i = 1 To Len(morse) - 1
        c = Mid(morse, i, 1)
        Select Case c
            Case ".":   extra = extra + ditLen + ditLen
                        addSound (ditLen)
                        addSilence (ditLen)
            Case "-":   extra = extra + dahlen + ditLen
                        addSound (dahlen)
                        addSilence (ditLen)
            Case " ":   extra = extra + charPause - ditLen
                        addSilence (charPause - ditLen + (extra * compression))
                        extra = 0
            Case "/":   extra = extra + wordPause - ditLen
                        addSilence (wordPause - ditLen + (extra * compression))
                        extra = 0
        End Select
        makeSample = sample
    Next i
    
End Function

Public Sub addSound(units As Integer)
Dim i As Integer
    'The values of the samples define the wave form, you can put here whatever
    'you fancy, any value between 0 and 255. Experiment with different values
    'to find the kind of sound that you find pleasant
    
    For i = 0 To units * timeunit - 1
        sample(sPos + 1) = 167 '0xA7
        sample(sPos + 2) = 129 '0x81
        sample(sPos + 3) = 167 '0xA7
        sample(sPos + 4) = 0
        sample(sPos + 5) = 89 '0x59
        sample(sPos + 6) = 127 '0x7F
        sample(sPos + 7) = 89 '0x59
        sample(sPos + 8) = 0
        sPos = sPos + 8
    Next i

End Sub

Public Sub addSilence(units As Integer)
Dim oPos As Long, i As Long

    oPos = sPos
    sPos = sPos + 8 * units * timeunit
    For i = oPos To sPos
        sample(i) = 0
    Next i

End Sub

Public Sub fill(a() As Byte, s As Integer, e As Integer, v As Byte)

    While (s < e)
        a(s) = v
        s = s + 1
    Wend
    
End Sub

Public Sub toMorse(sText As String)
Dim sbMorse As String, i As Long, j As Long
Dim c As String
    '  tr/a-z/A-Z/; #lowercase
    '  tr/ / /s; #sqeeze spaces
    '  s/^ *(.*?) *$/$1/; #chop start and end of ' '
    '  s# #~ #g; #mark word boundaries (with non-Morse character)
    '  s#([A-Z0-9.!,:?'`/()\"=+;_$@ -])#$toMorse{$1}.' '#eg; #put Morse in
    '  s#~#/#g; #re-mark word boundaries
    
    sbMorse = ""

    sText = UCase(sText)

        For i = 1 To Len(sText)
            c = Mid(sText, i, 1)
            
            If c = " " Then sbMorse = sbMorse + "  "
            
            If c <> " " Then
                For j = 0 To 52
                    If hMorse(j, 0) = c Then
                        sbMorse = sbMorse + hMorse(j, 1)
                        Exit For
                    End If
                Next j
                sbMorse = sbMorse + " "
            End If
        Next i
        
     Form1.txtMorse.Text = sbMorse
     
End Sub

Public Sub Initialise_MorseTables()

    hMorse(0, 0) = "A"
    hMorse(0, 1) = ".-"
    hMorse(1, 0) = "B"
    hMorse(1, 1) = "-..."
    hMorse(2, 0) = "C"
    hMorse(2, 1) = "-.-."
    hMorse(3, 0) = "D"
    hMorse(3, 1) = "-.."
    hMorse(4, 0) = "E"
    hMorse(4, 1) = "."
    hMorse(5, 0) = "F"
    hMorse(5, 1) = "..-."
    hMorse(6, 0) = "G"
    hMorse(6, 1) = "--."
    hMorse(7, 0) = "H"
    hMorse(7, 1) = "...."
    hMorse(8, 0) = "I"
    hMorse(8, 1) = ".."
    hMorse(9, 0) = "J"
    hMorse(9, 1) = ".---"
    hMorse(10, 0) = "K"
    hMorse(10, 1) = "-.-"
    hMorse(11, 0) = "L"
    hMorse(11, 1) = ".-.."
    hMorse(12, 0) = "M"
    hMorse(12, 1) = "--"
    hMorse(13, 0) = "N"
    hMorse(13, 1) = "-."
    hMorse(14, 0) = "O"
    hMorse(14, 1) = "---"
    hMorse(15, 0) = "P"
    hMorse(15, 1) = ".--."
    hMorse(16, 0) = "Q"
    hMorse(16, 1) = "--.-"
    hMorse(17, 0) = "R"
    hMorse(17, 1) = ".-."
    hMorse(18, 0) = "S"
    hMorse(18, 1) = "..."
    hMorse(19, 0) = "T"
    hMorse(19, 1) = "-"
    hMorse(20, 0) = "U"
    hMorse(20, 1) = "..-"
    hMorse(21, 0) = "V"
    hMorse(21, 1) = "...-"
    hMorse(22, 0) = "W"
    hMorse(22, 1) = ".--"
    hMorse(23, 0) = "X"
    hMorse(23, 1) = "-..-"
    hMorse(24, 0) = "Y"
    hMorse(24, 1) = "-.--"
    hMorse(25, 0) = "Z"
    hMorse(25, 1) = "--.."
    hMorse(26, 0) = "1"
    hMorse(26, 1) = ".----"
    hMorse(27, 0) = "2"
    hMorse(27, 1) = "..---"
    hMorse(28, 0) = "3"
    hMorse(28, 1) = "...--"
    hMorse(29, 0) = "4"
    hMorse(29, 1) = "....-"
    hMorse(30, 0) = "5"
    hMorse(30, 1) = "....."
    hMorse(31, 0) = "6"
    hMorse(31, 1) = "-...."
    hMorse(32, 0) = "7"
    hMorse(32, 1) = "--..."
    hMorse(33, 0) = "8"
    hMorse(33, 1) = "---.."
    hMorse(34, 0) = "9"
    hMorse(34, 1) = "----."
    hMorse(35, 0) = "0"
    hMorse(35, 1) = "-----"
    hMorse(36, 0) = "."
    hMorse(36, 1) = ".-.-.-"
    hMorse(37, 0) = "!"
    hMorse(37, 1) = "-.-.--"
    hMorse(38, 0) = ","
    hMorse(38, 1) = "--..--"
    hMorse(39, 0) = ":"
    hMorse(39, 1) = "---..."
    hMorse(40, 0) = "?"
    hMorse(40, 1) = "..--.."
    hMorse(41, 0) = "\"
    hMorse(41, 1) = ".----."
    hMorse(42, 0) = "`"
    hMorse(42, 1) = ".----." ' treat ` as "
    hMorse(43, 0) = "-"
    hMorse(43, 1) = "-....-"
    hMorse(44, 0) = "/"
    hMorse(44, 1) = "-..-."
    hMorse(45, 0) = "("
    hMorse(45, 1) = "-.--.-"
    hMorse(46, 0) = ")"
    hMorse(46, 1) = "-.--.-"
    hMorse(47, 0) = "="
    hMorse(47, 1) = "-...-"
    hMorse(48, 0) = "+"
    hMorse(48, 1) = ".-.-."
    hMorse(49, 0) = ";"
    hMorse(49, 1) = "-.-.-."
    hMorse(50, 0) = "_"
    hMorse(50, 1) = "..--.-"
    hMorse(51, 0) = "$"
    hMorse(51, 1) = "...-..-"
    hMorse(52, 0) = "@"
    hMorse(52, 1) = ".--.-."

End Sub

Public Sub PlayEnteredText()

makeSample Form1.txtMorse.Text

DSBD.lBufferBytes = UBound(sample)

Set DSB = DS.CreateSoundBuffer(DSBD, PCM)

DSB.WriteBuffer 0, 0, sample(0), DSBLOCK_ENTIREBUFFER

Form1.txtMorse.SelStart = 0
Form1.txtMorse.SelLength = 0
Form1.txtMorse.SetFocus

Form1.CursorPosTimer.Enabled = True

DSB.Play DSBPLAY_DEFAULT

End Sub

Public Sub PlayRandomText()
Dim i As Integer, j As Integer, RandomChar As String
Dim StartValue As Integer

Select Case PlayType                'choose whether to generate letters, numbers or
    Case LETTERS: StartValue = 0    'a combination of all possible characters
    Case NUMBERS: StartValue = 26
    Case ALL: StartValue = 0
End Select

While Not RandomTerminated
    For j = 1 To 5                  'produce 5 random words
        For i = 1 To 5              'of each 5 characters
            Randomize               'ensure that each character is really random
            RandomChar = hMorse(StartValue + Rnd * (PlayType - StartValue), 0)
            Form1.txtEnglish.Text = Form1.txtEnglish.Text + LCase(RandomChar)
        Next i
        Form1.txtEnglish.Text = Form1.txtEnglish.Text + " "
    Next j
    toMorse Form1.txtEnglish.Text   'convert the text to morse string
    PlayingCompleted = False
    PlayEnteredText
    DoEvents
    While Not PlayingCompleted      'we have to wait for the sequence to finish
        DoEvents                    'playing before generating the next one
        If RandomTerminated Then Exit Sub 'if terminated by clicking stop button
    Wend
    Form1.txtEnglish.Text = ""      'empty the text windows for the new sequence
    Form1.txtMorse.Text = ""
Wend

End Sub
