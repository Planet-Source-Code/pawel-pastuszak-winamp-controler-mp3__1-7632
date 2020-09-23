Attribute VB_Name = "WinAmpMod"
'WinAmpMod - winamp.bas
'BY: Pawel Pastuszak
'E-Mail: pastuszak@home.com
Option Explicit

'All the Declarations
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageCDS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long


Public Const WM_USER = &H400
Public Const WM_WA_IPC = WM_USER
Public Const WM_COPYDATA = &H4A
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const MF_BYCOMMAND = &H0
Public Const WU_MF_CHECKED = &H8

Public hwnd_winamp As Long
Public hMenuOptions As Long
Public hMenuWinamp As Long


Public Const IPC_DELETE = 101
Public Const IPC_ISPLAYING = 104
Public Const IPC_GETOUTPUTTIME = 105
Public Const IPC_JUMPTOTIME = 106
Public Const IPC_WRITEPLAYLIST = 120
Public Const IPC_SETPLAYLISTPOS = 121
Public Const IPC_SETVOLUME = 122
Public Const IPC_SETPANNING = 123
Public Const IPC_GETLISTLENGTH = 124
Public Const IPC_SETSKIN = 200
Public Const IPC_GETSKIN = 201
Public Const IPC_GETLISTPOS = 125
Public Const IPC_GETINFO = 126
Public Const IPC_GETEQDATA = 127
Public Const IPC_PLAYFILE = 100
Public Const IPC_GETPLAYLISTFILE = 211
Public Const IPC_CHDIR = 103

Public Const WINAMP_OPTIONS_EQ = 40036
Public Const WINAMP_OPTIONS_PLEDIT = 40040
Public Const WINAMP_VOLUMEUP = 40058
Public Const WINAMP_VOLUMEDOWN = 40059
Public Const WINAMP_FFWD5S = 40060
Public Const WINAMP_REW5S = 40061
Public Const WINAMP_BUTTON1 = 40044
Public Const WINAMP_BUTTON2 = 40045
Public Const WINAMP_BUTTON3 = 40046
Public Const WINAMP_BUTTON4 = 40047
Public Const WINAMP_BUTTON5 = 40048
Public Const WINAMP_BUTTON1_SHIFT = 40144
Public Const WINAMP_BUTTON4_SHIFT = 40147
Public Const WINAMP_BUTTON5_SHIFT = 40148
Public Const WINAMP_BUTTON1_CTRL = 40154
Public Const WINAMP_BUTTON2_CTRL = 40155
Public Const WINAMP_BUTTON5_CTRL = 40158
Public Const WINAMP_FILE_PLAY = 40029
Public Const WINAMP_OPTIONS_PREFS = 40012
Public Const WINAMP_OPTIONS_AOT = 40019
Public Const WINAMP_HELP_ABOUT = 40041          'About Window
Public Const WINAMP_FILE_QUIT = 40001           'Quit WinAmp
Public Const WINAMP_URL_WINDOW = 40155   'URL Window
Public Const WINAMP_FORWARD_TEN = 40195  'Forward 10 tracks
Public Const WINAMP_BACK_TEN = 40197     'Back 10 tracks
Public Const WINAMP_SHADEMODE = 40064    'Shape Mode

'Winamp menu item ID's (used to monitor items checked state)
Public Const WINAMP_MENU_SHUFFLE = &H9C57       'Shuffle Button
Public Const WINAMP_MENU_REPEAT = &H9C56        'Repeat Button
Public Const WINAMP_MENU_TIMEELAPSED = &H9C65
Public Const WINAMP_MENU_TIMEREMAINING = &H9C66
Public Const WINAMP_MENU_ALWAYSONTOP = &H9C53
Public Const WINAMP_MENU_DOUBLESIZE = &H9CE5
Public Const WINAMP_MENU_EASYMOVE = &H9CFA
'Winamp Command Definitions
'PLAYBACK
Public Const WINAMP_SHUFFLE = 40023        'Shuffle Button truns on/off
Public Const WINAMP_REPEAT = 40022         'Repeat Button truns on/off
Public Const WINAMP_DOUBLESIZE = 40165
'Windows
Public Const WINAMP_WINDMAIN = 40258
Public Const WINAMP_WINDBROWSER = 40298


Public Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As String
End Type

Public Type SONGTITLE
    sName As String
    sArtist As String
    sTrackNumber As String
End Type
Public TrackTitle As SONGTITLE


Public Function FindWinamp() As Long
'Find winamp window
'Returns 1 if winamp is open, 0 if not
    Dim hMenuSystem As Long
    
    hwnd_winamp = FindWindow("Winamp v1.x", vbNullString)
    hMenuSystem = GetSystemMenu(hwnd_winamp, 0)
    hMenuWinamp = GetSubMenu(hMenuSystem, 0)
    hMenuOptions = GetSubMenu(hMenuWinamp, 11)
    
    If hwnd_winamp Then FindWinamp = 1 Else FindWinamp = 0
End Function

Public Function DeletePlayList() As Long
'Clears the play list
    DeletePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_DELETE)
End Function
Public Function IsPlaying() As Long
'Returns:
'1 If playing
'3 if paused
'0 if stopped
    IsPlaying = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_ISPLAYING)
End Function

Public Function GetCurrentSongPosition() As Double
'Finds the current song position in milliseconds
    GetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETOUTPUTTIME)
End Function

Public Function GetSongLength() As Long
'Finds the song length in Seconds
    GetSongLength = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETOUTPUTTIME)
End Function

Public Function SetCurrentSongPosition(Optional Seconds As Long, Optional Ms As Long)
'Sets the current position in the song
'Returns:
'0 if success
'1 if eof
'-1 if not playing
    SetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, (Seconds * 1000 + Ms), IPC_JUMPTOTIME)
End Function


Public Function WritePlayList() As Long
'Writes the current playlist to C:\WINAMP_DIR\Winamp.m3u
'And then finds the play position
'Now obsolete, but good for old version of winamp
'Look at GetPlayListPosition
    WritePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_WRITEPLAYLIST)
End Function

Public Function SetPlayListPosition(Position As Integer) As Long
'Sets which song to play (0 being first)
    SetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, Position, IPC_SETPLAYLISTPOS)
End Function

Public Function SetVolume(Volume As Integer) As Long
'Sets the volume (Volume must be between 0 - 255)
    SetVolume = SendMessage(hwnd_winamp, WM_WA_IPC, Volume, IPC_SETVOLUME)
End Function

Public Function SetPanning(PanPosition As Integer) As Long
'Sets the panning (PanPosition must be between 0 - 255)
    SetPanning = SendMessage(hwnd_winamp, WM_WA_IPC, PanPosition, IPC_SETPANNING)
End Function

Public Function GetPlayListLength() As Long
'Gets amount of songs in play list
    GetPlayListLength = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTLENGTH)
End Function


Public Function GetPlayListPosition() As Long
'Returns which song its playing in the playlist
'0 being first
    GetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTPOS)
End Function

Public Function GetSamplerate() As Long
'Gets the samplerate
    GetSamplerate = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETINFO)
End Function

Public Function GetBitrate() As Long
'Gets the bitrate
    GetBitrate = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETINFO)
End Function

Public Function GetChannels() As Long
'Gets the channel
    GetChannels = SendMessage(hwnd_winamp, WM_WA_IPC, 2, IPC_GETINFO)
End Function

Public Function GetEQBandData(BandNumber As Integer) As Long
'Get each EQ banddata (0 being the first, 9 being last)
'Returns 0 - 255
    If BandNumber > 9 Then Exit Function
    GetEQBandData = SendMessage(hwnd_winamp, WM_WA_IPC, BandNumber, IPC_GETEQDATA)
End Function

Public Function GetEQPreampValue() As Long
'Gets the preamp value (Between 0 - 255)
    GetEQPreampValue = SendMessage(hwnd_winamp, WM_WA_IPC, 10, IPC_GETEQDATA)
End Function

Public Function GetEQEnabled()
'1 if EQ is enabled
'0 if it isn't
    GetEQEnabled = SendMessage(hwnd_winamp, WM_WA_IPC, 11, IPC_GETEQDATA)
End Function

Public Function GetEQAutoLoad()
'1 if EQ is autoloaded
'0 if it isn't
    GetEQAutoLoad = SendMessage(hwnd_winamp, WM_WA_IPC, 12, IPC_GETEQDATA)
End Function

Public Function PlayFile(FileToPlay As String) As Long
'Adds FileToPlay to the play list
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_PLAYFILE
    CDS.lpData = FileToPlay
    CDS.cbData = Len(FileToPlay) + 1
    PlayFile = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function ChangeDirectory(Directory As String) As Long
'Changes directory
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_CHDIR
    CDS.lpData = Directory
    CDS.cbData = Len(Directory) + 1
    ChangeDirectory = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function ToggleEQWindow() As Long
'Turns on or off the EQ window
    ToggleEQWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_EQ, 0)
End Function

Public Function TogglePlayListWindow() As Long
'Turns on or off play list window
    TogglePlayListWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PLEDIT, 0)
End Function

Public Function VolumeUp() As Long
'Raises the volume a tiny bit
    VolumeUp = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEUP, 0)
End Function
Public Function VolumeDown() As Long
'Sets the volume down a tiny bit
    VolumeDown = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEDOWN, 0)
End Function

Public Function Rewind() As Long
'Rewinds by 5 seconds
    Rewind = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_REW5S, 0)
End Function

Public Function FastForward() As Long
'Fast forwards by 5 seconds
    FastForward = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_FFWD5S, 0)
End Function

Public Function PreviousTrack() As Long
'Plays the previous song
    PreviousTrack = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1, 0)
End Function

Public Function PlaySong() As Long
'Plays the current song
    PlaySong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2, 0)
End Function

Public Function PauseSong() As Long
'Pauses playing
    PauseSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON3, 0)
End Function
Public Function StopSong() As Long
'Stops playing
    StopSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4, 0)
End Function

Public Function NextTrack() As Long
'Plays the next song in the playlist
    NextTrack = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5, 0)
End Function

Public Function FadeStop() As Long
'slowly fades away until it stops
    FadeStop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4_SHIFT, 0)
End Function

Public Function FirstSong() As Long
'Goes to the first song in the play list
    FirstSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1_CTRL, 0)
End Function

Public Function LastSong() As Long
'Goes to the last song in the play list
    LastSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5_CTRL, 0)
End Function
Public Function OpenLocation() As Long
'Shows Open Location Dialog
    OpenLocation = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2_CTRL, 0)
End Function
Public Function LoadFile() As Long
'Shows Load a file dialog
    LoadFile = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_FILE_PLAY, 0)
End Function
Public Function ShowPreferences() As Long
'Shows Preferences Dialog
    ShowPreferences = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PREFS, 0)
End Function

Public Function ToggleAlwaysOnTop() As Long
'Turns Always On Top On and Off
    ToggleAlwaysOnTop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_AOT, 0)
End Function

Public Function ShowAbout() As Long
'Shows About Box
    ShowAbout = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_HELP_ABOUT, 0)
End Function
Public Function QuitWinamp() As Long
'Quit Winamp
    QuitWinamp = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_FILE_QUIT, 0)
End Function
'gets the current song playing
Public Function GetSongTitle() As String
Dim sSong_buffer As String * 100
'get the title
Call GetWindowText(hwnd_winamp, sSong_buffer, 99)
'edit the sSong_buffer
'xx.aaaaaa - tttttttt - Winamp
'get the track number
Dim iTrackNumber As Integer
iTrackNumber = InStr(1, sSong_buffer, ".")
TrackTitle.sTrackNumber = Trim$(Left(sSong_buffer, Val(iTrackNumber - 1)))
'get the artist
Dim iArtist As Integer
iArtist = InStr(iTrackNumber, sSong_buffer, "-")
TrackTitle.sArtist = Trim(Mid(sSong_buffer, iTrackNumber + 1, (iArtist - iTrackNumber - 1)))
Dim iSong As Integer
iSong = InStr(iArtist + 1, sSong_buffer, "-")
TrackTitle.sName = Trim(Mid(sSong_buffer, iArtist + 1, (iSong - iArtist - 1)))
'get the title
GetSongTitle = TrackTitle.sTrackNumber + ". " + TrackTitle.sArtist + " - " + TrackTitle.sName

End Function
Public Function IsShuffle() As Boolean
'check if the Shuffle button was clicked
    If (GetMenuState(hMenuOptions, WINAMP_MENU_SHUFFLE, MF_BYCOMMAND)) = WU_MF_CHECKED Then
        IsShuffle = True
    Else
        IsShuffle = False
    End If
End Function

Public Function IsRepeat() As Boolean
'check if the Repeat Button was click or not
    If (GetMenuState(hMenuOptions, WINAMP_MENU_REPEAT, MF_BYCOMMAND) = WU_MF_CHECKED) Then
        IsRepeat = True
    Else
        IsRepeat = False
    End If
End Function
Public Function Back10Songs() As Long
'Goes to the first song in the play list
    Back10Songs = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BACK_TEN, 0)
End Function

Public Function Forward10Songs() As Long
'Goes to the last song in the play list
    Forward10Songs = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_FORWARD_TEN, 0)
End Function

Public Function ShowURLWindow() As Long
'show's the url window
    ShowURLWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_URL_WINDOW, 0)
End Function
Public Function ToggleShuffle() As Long
'check off and on the Shuffle Button
ToggleShuffle = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_SHUFFLE, 0)
End Function
Public Function ToggleRepeat() As Long
'check off the Repeat on/off button
ToggleRepeat = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_REPEAT, 0)
End Function
Public Function ToggleMain() As Long
'trun on/off the main window
ToggleMain = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_WINDMAIN, 0)
End Function
Public Function IsTimeElapsed() As Boolean
    If (GetMenuState(hMenuOptions, WINAMP_MENU_TIMEELAPSED, MF_BYCOMMAND) = WU_MF_CHECKED) Then
        IsTimeElapsed = True
    Else
        IsTimeElapsed = False
    End If
End Function


Public Function IsTimeRemaining() As Boolean
    If (GetMenuState(hMenuOptions, WINAMP_MENU_TIMEREMAINING, MF_BYCOMMAND) = WU_MF_CHECKED) Then
        IsTimeRemaining = True
    Else
        IsTimeRemaining = False
    End If

End Function
Public Function IsAlwaysOnTop() As Boolean
    If (GetMenuState(hMenuOptions, WINAMP_MENU_ALWAYSONTOP, MF_BYCOMMAND) = WU_MF_CHECKED) Then
        IsAlwaysOnTop = True
    Else
        IsAlwaysOnTop = False
    End If
End Function

Public Function IsDoubleSize() As Boolean
    If (GetMenuState(hMenuOptions, WINAMP_MENU_DOUBLESIZE, MF_BYCOMMAND) = WU_MF_CHECKED) Then
        IsDoubleSize = True
    Else
        IsDoubleSize = False
    End If
End Function

Public Function IsEasyMove() As Boolean
    If (GetMenuState(hMenuOptions, WINAMP_MENU_EASYMOVE, MF_BYCOMMAND) = WU_MF_CHECKED) Then
        IsEasyMove = True
    Else
        IsEasyMove = False
    End If
End Function
Public Function ToggleDoubleSize() As Long
    ToggleDoubleSize = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_DOUBLESIZE, 0)
End Function
Public Function ToggleBrowser() As Long
    ToggleBrowser = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_WINDBROWSER, 0)
End Function
Public Function ToggleShade() As Long
    ToggleShade = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_SHADEMODE, 0)
End Function
Public Function ConvertTime(TimeSeconds As Long)
    Dim TimeMin         'Time in minutes & decimal      (5.911 minutes)
    Dim TimeMinOnly     'Time in whole minutes only     (5.0 minutes)
    Dim TimeSecOnly     'Seconds portion of time only   (54 seconds (.911 * 60))
    
    Dim MinutesStr      'Minutes converted to a string  ("05" minutes)
    Dim SecondsStr      'Seconds converted to a string  ("54" seconds)
    
        TimeMin = (TimeSeconds / 60)
        TimeMinOnly = Int(TimeMin)
        TimeSecOnly = Int((TimeMin - TimeMinOnly) * 60)
        
        If TimeMinOnly < 10 Then
            MinutesStr = "0" & TimeMinOnly
        Else
            MinutesStr = TimeMinOnly
        End If
        
        If TimeSecOnly < 10 Then
            SecondsStr = "0" & TimeSecOnly
        Else
            SecondsStr = TimeSecOnly
        End If

        ConvertTime = MinutesStr & ":" & SecondsStr
End Function

