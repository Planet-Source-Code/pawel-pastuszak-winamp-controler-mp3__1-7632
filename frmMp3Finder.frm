VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Winamp controler: By Pawel Pastuszak"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   38
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSavePlayList 
      Caption         =   "Save PL"
      Height          =   375
      Left            =   2760
      TabIndex        =   33
      ToolTipText     =   "Save PlayList to Defualt m3u file"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdToggleShade 
      Caption         =   "Toggle Shade"
      Height          =   375
      Left            =   5400
      TabIndex        =   32
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdToggleBrowser 
      Caption         =   "Toggle MB"
      Height          =   375
      Left            =   4080
      TabIndex        =   29
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdToggleEQ 
      Caption         =   "Toggle EQ"
      Height          =   375
      Left            =   4080
      TabIndex        =   28
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdTogglePlayList 
      Caption         =   "Toggle PL"
      Height          =   375
      Left            =   4080
      TabIndex        =   27
      ToolTipText     =   "Toggle Playlist"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdToggleMain 
      Caption         =   "Toggle Main"
      Height          =   375
      Left            =   4080
      TabIndex        =   26
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdVolumeDown 
      Caption         =   "Volume Down-"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdVolumeUp 
      Caption         =   "Volume Up+"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdShutDown 
      Caption         =   "Shut Down"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdLastSong 
      Caption         =   "Last Song"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdForwardTenSongs 
      Caption         =   "Forward 10+"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackTenSongs 
      Caption         =   "Back 10-"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirstSong 
      Caption         =   "First Song"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3600
      Top             =   4320
   End
   Begin VB.CommandButton cmdAboutWinAmp 
      Caption         =   "&WinAmp About"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "&Start WinAmp"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuitWinamp 
      Caption         =   "&Quit Winamp"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   7815
   End
   Begin VB.Frame fraPlayControls 
      Caption         =   "Controls"
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6495
      Begin VB.HScrollBar hBalance 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   4680
         Max             =   127
         Min             =   -127
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin VB.HScrollBar hVolume 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   3000
         Max             =   255
         TabIndex        =   18
         Top             =   1800
         Value           =   255
         Width           =   1575
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "&Pause"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   100
         TabIndex        =   0
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "&Play"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblKHZ 
         Alignment       =   1  'Right Justify
         Caption         =   "0 KHZ"
         Height          =   255
         Left            =   5520
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblKBIT 
         Alignment       =   1  'Right Justify
         Caption         =   "0 KBIT"
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblChannels 
         Caption         =   "Channels:"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTrackNumber 
         Caption         =   "Track #:"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total: 00:00"
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "00:00"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance:"
         Height          =   255
         Left            =   4680
         TabIndex        =   25
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblDoubleSize 
         Caption         =   "Double Size: Off"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblOnTop 
         Caption         =   "On Top: Off"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblVolume 
         Caption         =   "Volume:"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblRepeat 
         Caption         =   "Repeat: Off"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblShuffle 
         Caption         =   "Shuffle: Off"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblTrack 
         Caption         =   "No Track"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By: Pawel Pastuszak

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long


Private Sub cmdAboutWinAmp_Click()
Call ShowAbout
End Sub

Private Sub cmdBack_Click()
Call PreviousTrack 'go back a track
End Sub

Private Sub cmdBackTenSongs_Click()
Call Back10Songs
End Sub

Private Sub cmdFirstSong_Click()
Call FirstSong
End Sub

Private Sub cmdLastSong_Click()
Call LastSong
End Sub

Private Sub cmdNext_Click()
'next track/song
Call NextTrack
End Sub

Private Sub cmdPause_Click()
Call PauseSong          'pause the current song

End Sub

Private Sub cmdPlay_Click()
Call PlaySong       'Play winamp song.
End Sub

Private Sub cmdQuitWinamp_Click()
Call QuitWinamp 'Quit Winamp
End Sub

Private Sub cmdSavePlayList_Click()
WinAmpMod.WritePlayList
End Sub

Private Sub CmdStart_Click()
x = FindWinamp
If x = 1 Then
    'found the Mp3Player
    List1.AddItem "Found Mp3 Player"
Else
    'not found
    List1.AddItem x & " Not found"
End If

x = IsPlaying
If x = 1 Then
    List1.AddItem "Playing sound"
ElseIf x = 3 Then
    List1.AddItem "Pasued Song"
Else
    List1.AddItem "Stoped Song"
End If

List1.AddItem "Songs in the Play List: " & GetPlayListLength
List1.AddItem "Play List Position: " & (GetPlayListPosition + 1)
List1.AddItem "Current Song Position: " & GetCurrentSongPosition
SongPositionSecond = (GetCurrentSongPosition / 1000)
List1.AddItem "Currently at " + CStr(SongPositionSecond) + " seconds with a song length of " + CStr(GetSongLength) + " seconds."

List1.AddItem FindWindow("Netscape v1.0x", vbNullString)

End Sub


Private Sub cmdStop_Click()
Call StopSong       'Stop A Song
End Sub


Private Sub cmdToggleBrowser_Click()
Call ToggleBrowser
End Sub

Private Sub cmdToggleEQ_Click()
Call ToggleEQWindow
End Sub

Private Sub cmdToggleMain_Click()
Call ToggleMain
End Sub

Private Sub cmdTogglePlayList_Click()
Call TogglePlayListWindow
End Sub

Private Sub cmdToggleShade_Click()
Call ToggleShade
End Sub

Private Sub cmdVolumeDown_Click()
Call VolumeDown
End Sub

Private Sub cmdVolumeUp_Click()
Call VolumeUp
End Sub

Private Sub Command1_Click()
Call SetCursorPos(0, 222)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Me.Caption = KeyCode
Select Case KeyCode
    Case 45:    'Key -> 0 on Null Pad
        If WinAmpMod.IsPlaying = 1 Then
            cmdPause_Click
        Else
            cmdPlay_Click
        End If
    Case 13:
        Call cmdStop_Click   'stop playing the music.
    Case 46:
        cmdPause_Click 'pasued the music
    Case 37:
        cmdBack_Click   'next song
    Case 39:
        cmdNext_Click   'back a song
    Case 38:
        cmdVolumeUp_Click   'volume up
    Case 40:
        cmdVolumeDown_Click 'volume down
    Case 35:
        cmdBackTenSongs_Click 'back 10 songs
    Case 34:
        cmdForwardTenSongs_Click 'skip ever ten songs
    Case 33:
        Call WinAmpMod.FastForward
    Case 36:
        Call WinAmpMod.Rewind
    Case 12:
        Call cmdPlay_Click  'resart muscic
End Select
'Me.Caption = Str(KeyCode)
End Sub

Private Sub Form_Load()
'Find the Mp3 Player
Mp3Player = FindWinamp
If Mp3Player = 1 Then
    'found the Mp3Player
    List1.AddItem "Found Mp3 Player"
Else
    'not found
    List1.AddItem x & " Not found"
End If

End Sub


Private Sub hBalance_Change()
If hBalance.Value = 0 Then
    lblBalance.Caption = "Balance: CENTER"
ElseIf hBalance.Value < 0 Then
    lblBalance.Caption = "Balance: LEFT"
Else
    lblBalance.Caption = "Balance: RIGHT"
End If
Call WinAmpMod.SetPanning(hBalance.Value)
End Sub

Private Sub hBalance_Scroll()
If hBalance.Value = 0 Then
    lblBalance.Caption = "Balance: CENTER"
ElseIf hBalance.Value < 0 Then
    lblBalance.Caption = "Balance: LEFT"
Else
    lblBalance.Caption = "Balance: RIGHT"
End If
Call WinAmpMod.SetPanning(hBalance.Value)

End Sub

Private Sub hVolume_Change()
Call SetVolume(hVolume.Value)
End Sub

Private Sub hVolume_Scroll()
Call SetVolume(hVolume.Value)
End Sub


Private Sub hVolume_Validate(Cancel As Boolean)
'Call SetVolume(hVolume.Value)
End Sub

Private Sub lblDoubleSize_Click()
Call ToggleDoubleSize
End Sub

Private Sub cmdForwardTenSongs_Click()
Call Forward10Songs
End Sub

Private Sub lblOnTop_Click()
Call ToggleAlwaysOnTop
End Sub

Private Sub lblRepeat_Click()
Call ToggleRepeat
End Sub

Private Sub lblShuffle_Click()
Call ToggleShuffle
End Sub

Private Sub Timer1_Timer()
If FindWinamp = 0 Then Exit Sub
If IsShuffle = True Then
    lblShuffle.Caption = "Shuffle: On"
Else
    lblShuffle.Caption = "Shuffle: Off"
End If
If IsRepeat = True Then
    lblRepeat.Caption = "Repeat: On"
Else
    lblRepeat.Caption = "Repeat: Off"
End If
If IsAlwaysOnTop = True Then
    lblOnTop.Caption = "On Top: On"
Else
    lblOnTop.Caption = "On Top: Off"
End If

If IsDoubleSize = True Then
    lblDoubleSize.Caption = "Double Size: On"
Else
    lblDoubleSize.Caption = "Double Size: Off"
End If
'time
If WinAmpMod.IsTimeElapsed = True Then
    lblTime.Caption = ConvertTime(GetCurrentSongPosition / 1000)
Else
    Dim p As Long
    Dim l As Long
    p = (GetCurrentSongPosition / 1000)
    l = GetSongLeght
    lblTime.Caption = ConvertTime(l - p)
End If
lblTotal.Caption = ConvertTime(GetSongLength)

lblTrack.Caption = GetSongTitle
If IsPlaying = 0 Then
    lblTrack.Caption = lblTrack.Caption + " [Stopped]"
ElseIf IsPlaying = 3 Then
    lblTrack.Caption = lblTrack.Caption + " [Paused]"
End If
'get the curent track and how many tracks they are
lblTrackNumber.Caption = "Track #: " + Str(GetPlayListPosition + 1) + " of " + Str(GetPlayListLength)
'get the number of channels
lblChannels.Caption = "Channels: " + Str(WinAmpMod.GetChannels)
'get the bitrate.
lblKBIT.Caption = Str(GetBitrate) + " KBIT/s"
'get the KHZ
lblKHZ.Caption = Str(WinAmpMod.GetSamplerate) + " KHZ"
'Randomize
'x = Int(1 * Rnd * 420)
'y = Int(1 * Rnd * 420)
'Call SetCursorPos(x, y)

End Sub
