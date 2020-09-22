VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VWA Game"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer TimerFlamme 
      Enabled         =   0   'False
      Interval        =   255
      Left            =   6600
      Top             =   4800
   End
   Begin VB.Timer TimerPapier 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   7080
      Top             =   4800
   End
   Begin VB.Timer TimerSoundOff 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8520
      Top             =   4800
   End
   Begin VB.Timer TimerGameTime 
      Left            =   8040
      Top             =   4800
   End
   Begin VB.Timer TimerEinpennen 
      Left            =   7560
      Top             =   4800
   End
   Begin VB.Image ImagePFlames 
      Height          =   930
      Left            =   5160
      Picture         =   "FormMain.frx":0000
      Top             =   3960
      Width           =   900
   End
   Begin VB.Image ImagePKorb 
      Height          =   990
      Left            =   5160
      Picture         =   "FormMain.frx":078F
      Top             =   4200
      Width           =   900
   End
   Begin VB.Image ImagePapier 
      Height          =   495
      Left            =   3720
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image ImageDozent 
      Height          =   1800
      Left            =   2400
      Picture         =   "FormMain.frx":0D56
      Stretch         =   -1  'True
      Top             =   4140
      Width           =   1605
   End
   Begin VB.Label Label4 
      Caption         =   $"FormMain.frx":12F0
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   9015
   End
   Begin VB.Label LabelScore 
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label LabelPunkte 
      BackStyle       =   0  'Transparent
      Caption         =   "Punkte:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label LabelPenner 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   4360
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Schon mal eingeschlafen:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label LabelTime 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label LabelLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   5400
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   39
      Left            =   7080
      Picture         =   "FormMain.frx":1448
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   38
      Left            =   6480
      Picture         =   "FormMain.frx":16AE
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   37
      Left            =   5760
      Picture         =   "FormMain.frx":1914
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   36
      Left            =   5160
      Picture         =   "FormMain.frx":1B7A
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   35
      Left            =   4440
      Picture         =   "FormMain.frx":1DE0
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   34
      Left            =   3720
      Picture         =   "FormMain.frx":2046
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   33
      Left            =   3000
      Picture         =   "FormMain.frx":22AC
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   465
      Index           =   32
      Left            =   2280
      Picture         =   "FormMain.frx":2512
      Stretch         =   -1  'True
      Top             =   940
      Width           =   375
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   31
      Left            =   7320
      Picture         =   "FormMain.frx":2778
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   30
      Left            =   6600
      Picture         =   "FormMain.frx":29DE
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   29
      Left            =   5880
      Picture         =   "FormMain.frx":2C44
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   28
      Left            =   5040
      Picture         =   "FormMain.frx":2EAA
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   27
      Left            =   4320
      Picture         =   "FormMain.frx":3110
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   26
      Left            =   3600
      Picture         =   "FormMain.frx":3376
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   25
      Left            =   2760
      Picture         =   "FormMain.frx":35DC
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   585
      Index           =   24
      Left            =   1920
      Picture         =   "FormMain.frx":3842
      Stretch         =   -1  'True
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   23
      Left            =   7560
      Picture         =   "FormMain.frx":3AA8
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   22
      Left            =   6600
      Picture         =   "FormMain.frx":3D0E
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   21
      Left            =   5760
      Picture         =   "FormMain.frx":3F74
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   20
      Left            =   4920
      Picture         =   "FormMain.frx":41DA
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   19
      Left            =   4080
      Picture         =   "FormMain.frx":4440
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   18
      Left            =   3240
      Picture         =   "FormMain.frx":46A6
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   17
      Left            =   2280
      Picture         =   "FormMain.frx":490C
      Stretch         =   -1  'True
      Top             =   1720
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   705
      Index           =   16
      Left            =   1320
      Picture         =   "FormMain.frx":4B72
      Stretch         =   -1  'True
      Top             =   1725
      Width           =   615
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   15
      Left            =   7800
      Picture         =   "FormMain.frx":4DD8
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   14
      Left            =   6840
      Picture         =   "FormMain.frx":503E
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   13
      Left            =   5880
      Picture         =   "FormMain.frx":52A4
      Stretch         =   -1  'True
      Top             =   2230
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   12
      Left            =   4800
      Picture         =   "FormMain.frx":550A
      Stretch         =   -1  'True
      Top             =   2230
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   11
      Left            =   3840
      Picture         =   "FormMain.frx":5770
      Stretch         =   -1  'True
      Top             =   2230
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   10
      Left            =   2760
      Picture         =   "FormMain.frx":59D6
      Stretch         =   -1  'True
      Top             =   2240
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   9
      Left            =   1800
      Picture         =   "FormMain.frx":5C3C
      Stretch         =   -1  'True
      Top             =   2240
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   825
      Index           =   8
      Left            =   720
      Picture         =   "FormMain.frx":5EA2
      Stretch         =   -1  'True
      Top             =   2240
      Width           =   735
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   7
      Left            =   7920
      Picture         =   "FormMain.frx":6108
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   6
      Left            =   6840
      Picture         =   "FormMain.frx":636E
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   5
      Left            =   5760
      Picture         =   "FormMain.frx":65D4
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   4
      Left            =   4800
      Picture         =   "FormMain.frx":683A
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   3
      Left            =   3720
      Picture         =   "FormMain.frx":6AA0
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   2
      Left            =   2640
      Picture         =   "FormMain.frx":6D06
      Stretch         =   -1  'True
      Top             =   2780
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   1
      Left            =   1560
      Picture         =   "FormMain.frx":6F6C
      Stretch         =   -1  'True
      Top             =   2780
      Width           =   855
   End
   Begin VB.Image ImageHoerer 
      Height          =   945
      Index           =   0
      Left            =   480
      Picture         =   "FormMain.frx":71D2
      Stretch         =   -1  'True
      Top             =   2780
      Width           =   855
   End
   Begin VB.Image ImageRoom 
      Height          =   5820
      Left            =   120
      Picture         =   "FormMain.frx":7438
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9000
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'*** VWA Hoerergame
'********************
'*** Carsten Skalla
'*** Samstag, 25.10.2003
'*** Das Soundmodul ist ausm Web (planetsourcecode.com)
'*****************************************************************
Option Explicit


Dim HoererStat(39) As Byte
Dim pennscore As Long
Dim xd As Integer, yd As Integer
Dim xh As Integer, yh As Integer
Dim np As Integer
Dim FlammenImKorb As Byte
Dim picDummyFlame(3) As StdPicture
Dim picDummyHoerer(7) As StdPicture
Dim picDummyPapier(8) As StdPicture

Const maxhoerer = 39
Const pennmax = 7
Const pennergewecktpunkte = 1
Const levelwert = 10
Const maxleveldauer = 20
Const einpennspeed = 1000



Private Sub Form_Activate()
Dim ant As String
ant = CStr(MsgBox("Geht gleich los...", vbOKOnly, "Bereit?"))
TimerSoundOff.Enabled = False
TimerEinpennen.Enabled = True
TimerGameTime.Enabled = True
End Sub

'hoerer 47  mal 66 pixel gifs
'room 600 mal 388 pixel bmp

Private Sub Form_Load()
Randomize
FlammenImKorb = 1
pennscore = 0
ImagePFlames.Visible = False
ImageRoom.Picture = LoadPicture(App.Path & "\room.bmp")
ImageDozent.Picture = LoadPicture(App.Path & "\dozent.gif")
TimerSoundOff.Enabled = False
TimerEinpennen.Enabled = False
TimerGameTime.Enabled = False
Call InitGrafik
TimerEinpennen.Interval = einpennspeed
TimerGameTime.Interval = 1000
LabelTime.Caption = maxleveldauer
LabelLevel.Caption = "1"
End Sub

Private Sub InitGrafik()
'hoerer init / anordnen...
Dim i As Long
'dummys füllen
For i = 0 To 3
    Set picDummyFlame(i) = LoadPicture(App.Path & "\papierkorbflammen" & CStr(i) & ".gif")
Next
For i = 0 To 7
    Set picDummyHoerer(i) = LoadPicture(App.Path & "\hoerer" & CStr(i) & ".gif")
Next
For i = 0 To 8
    Set picDummyPapier(i) = LoadPicture(App.Path & "\papier" & CStr(i) & ".gif")
Next
For i = maxhoerer To 0 Step -1
    ImageHoerer(i).Picture = picDummyHoerer(1)
    ImageHoerer(i).ZOrder 0
    HoererStat(i) = 1
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
TimerEinpennen.Interval = 0
TimerGameTime.Interval = 0
TimerSoundOff.Enabled = False
TimerEinpennen.Enabled = False
TimerGameTime.Enabled = False
End Sub

Private Sub ImageFlameDummy_Click(Index As Integer)
End Sub

Private Sub ImageHoerer_Click(Index As Integer)
'hoerer aufwecken
Dim n As Integer
If HoererStat(Index) > 1 Then
    n = Rnd(1) * 80
    If HoererStat(Index) > 5 Then
        Call PapierFlug(Index)
        n = 0
        HoererStat(Index) = 0
        ImageHoerer(Index).Picture = picDummyHoerer(0)
    Else
        HoererStat(Index) = 1
        ImageHoerer(Index).Picture = picDummyHoerer(1)
    End If
    
    pennscore = pennscore + pennergewecktpunkte
    Call PrintScore
    
    Select Case n
        Case 0: 'nix
        Case 1: PlayLargeSound App.Path & "\ow.wav", WaveFiles, "sfx"
        Case 2: PlayLargeSound App.Path & "\muede.wav", WaveFiles, "sfx"
        Case 3: PlayLargeSound App.Path & "\er.wav", WaveFiles, "sfx"
        Case 4: PlayLargeSound App.Path & "\merke.wav", WaveFiles, "sfx"
        Case 5: PlayLargeSound App.Path & "\watte.wav", WaveFiles, "sfx"
        Case 6: PlayLargeSound App.Path & "\zeit.wav", WaveFiles, "sfx"
        Case 7: PlayLargeSound App.Path & "\klar.wav", WaveFiles, "sfx"
        Case 8: PlayLargeSound App.Path & "\danke.wav", WaveFiles, "sfx"
        Case 9: PlayLargeSound App.Path & "\bloed.wav", WaveFiles, "sfx"
        Case 10: PlayLargeSound App.Path & "\oops.wav", WaveFiles, "sfx"
        Case Else: PlayLargeSound App.Path & "\hit.wav", WaveFiles, "sfx"
    End Select
    If n <> 0 Then TimerSoundOff.Enabled = True
    
    Else
    PlayLargeSound App.Path & "\yes.wav", WaveFiles, "sfx"
    TimerSoundOff.Enabled = True
    End If
End Sub

Private Sub PapierFlug(hoerer As Integer)
xd = ImageDozent.Left + 1300
yd = ImageDozent.Top + 500

xh = (ImageHoerer(hoerer).Left - xd + 400) / 8
yh = (ImageHoerer(hoerer).Top - yd) / 8
ImagePapier.ZOrder 0
np = 1
TimerPapier.Enabled = True
End Sub

Private Sub TimerEinpennen_Timer()
'einen hoerer pennen lassen
Dim n As Integer
Dim gonePennen As Boolean
gonePennen = False
Do
n = Rnd(maxhoerer) * maxhoerer
If HoererStat(n) < pennmax Then
    gonePennen = True
    HoererStat(n) = HoererStat(n) + 1
    ImageHoerer(n).Picture = picDummyHoerer(HoererStat(n))
    'ImageHoerer(n).Refresh
    DoEvents
    If HoererStat(n) = pennmax Then
        LabelPenner.Caption = CInt(LabelPenner.Caption) + 1
        End If
    End If
DoEvents
Loop Until gonePennen
If CInt(LabelPenner.Caption) = (maxhoerer + 1) / 2 Then
    Call Verloren
    End If
End Sub


Private Sub TimerGameTime_Timer()
'ein level weiterschalten.
'nun gehts fixer mit dem einpennen...
LabelTime.Caption = CInt(LabelTime.Caption) - 1
If CInt(LabelTime.Caption) = -1 Then
    LabelLevel.Caption = CInt(LabelLevel.Caption) + 1
    TimerEinpennen.Interval = einpennspeed / CInt(LabelLevel.Caption) * 2
    LabelTime.Caption = CStr(maxleveldauer)
    'oops sound weil neuer level
    PlayLargeSound App.Path & "\perfect.wav", WaveFiles, "sfx"
    TimerSoundOff.Enabled = True
    Call Initlevel
    If CInt(LabelLevel.Caption) = 10 Then
        ImagePFlames.Visible = True
        TimerFlamme.Enabled = True
        End If
    End If
End Sub

Private Sub Initlevel()
'neuer level!
Dim n As Long
Dim i As Long
Dim m As Long
pennscore = pennscore + levelwert
Call PrintScore
For i = maxhoerer To 0 Step -1
    If HoererStat(i) <= 5 Then
        HoererStat(i) = HoererStat(i) + 1
        ImageHoerer(n).Picture = picDummyHoerer(HoererStat(i))
        End If
Next
End Sub

Private Sub PrintScore()
'score aktualisieren
LabelScore.Caption = Right("00000" & CStr(pennscore), 5)
LabelScore.Refresh
End Sub

Private Sub Verloren()
'alle eingepennt!
TimerEinpennen.Interval = 0
TimerGameTime.Interval = 0
TimerSoundOff.Enabled = False
TimerEinpennen.Enabled = False
TimerGameTime.Enabled = False
    PlayLargeSound App.Path & "\bier.wav", WaveFiles, "sfx"
    TimerSoundOff.Enabled = True
MsgBox ("Verloren! " & ((maxhoerer + 1) / 2) & " mal ist ein Hörer" & vbNewLine & _
        "tief eingeschlafen und hat geschnarcht!" & vbNewLine & _
        vbNewLine & "Ihr Score: " & pennscore)
TimerFlamme.Enabled = False
End
End Sub

Private Sub TimerPapier_Timer()
'papier fliegt...
ImagePapier.Visible = False
ImagePapier.Picture = picDummyPapier(np)
ImagePapier.Left = xd + xh * np - ImagePapier.Width / 2
ImagePapier.Top = yd + yh * np - ImagePapier.Height / 2
ImagePapier.Visible = True
DoEvents
np = np + 1
If np = 9 Then
    PlayLargeSound App.Path & "\ow.wav", WaveFiles, "sfx"
    TimerSoundOff.Enabled = True
    ImagePapier.Visible = False
    TimerPapier.Enabled = False
    End If
End Sub

Private Sub TimerSoundOff_Timer()
'sound wieder off
  If PersendagePlayedOfAnOpenedSound("sfx") = 100 Then
    StopLargeSound "sfx"
    TimerSoundOff.Enabled = False
  End If
End Sub

Private Sub TimerFlamme_Timer()
'flamme umbasteln
FlammenImKorb = FlammenImKorb + 1
If FlammenImKorb = 4 Then FlammenImKorb = 1
    ImagePFlames.Picture = picDummyFlame(FlammenImKorb)
End Sub

