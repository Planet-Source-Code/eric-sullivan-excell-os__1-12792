VERSION 5.00
Begin VB.Form frmMed 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MacSHELL Media Player"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2760
   End
   Begin VB.HScrollBar Volume 
      Height          =   330
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox TimeWindow 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "dasfvdsfv"
      ToolTipText     =   "Time"
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image Image8 
      Height          =   390
      Left            =   3480
      Picture         =   "frmMed.frx":0000
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Image Image7 
      Height          =   390
      Left            =   2400
      Picture         =   "frmMed.frx":079E
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   3480
      Picture         =   "frmMed.frx":0F14
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   1320
      Picture         =   "frmMed.frx":1677
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   240
      Picture         =   "frmMed.frx":1DF2
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   1320
      Picture         =   "frmMed.frx":257E
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   2400
      Picture         =   "frmMed.frx":2CE3
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   240
      Picture         =   "frmMed.frx":3437
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label TrackTime 
      BackColor       =   &H00000000&
      Caption         =   "dfgfgn"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label TotalTrack 
      BackColor       =   &H00000000&
      Caption         =   "fgnfgn"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   240
      Top             =   480
      Width           =   4215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   3060
      Left            =   110
      Top             =   110
      Width           =   4620
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3120
      Left            =   75
      Top             =   75
      Width           =   4680
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   3165
      Left            =   45
      Top             =   45
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3240
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ExcellOS - CD Player"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   120
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3255
      Left            =   4800
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FastForwardSpeed As Long        ' seconds to seek for ff/rew
Dim Playing As Boolean                ' true if CD is currently playing
Dim CDLoad As Boolean                  ' true if CD is the the player
Dim TotalTracks As Integer              ' total tracks tracks on audio CD
Dim TrackLength() As String              ' array containing length of each track
Dim Track As Integer                     ' current track
Dim Minute As Integer                   ' current minute on track
Dim Second As Integer                  ' current second on track
Dim Command As String                 ' string to hold mci command strings
Dim hmixer As Long                   ' mixer handle
Dim volCtrl As MIXERCONTROL         ' Waveout volume control.
Dim CDROMopen As Boolean

Private Function SendMCIString(Cmd As String, fShowError As Boolean) As Boolean
    Static rc As Long               'return code
    Static errStr As String * 400
    
    rc = mciSendString(Cmd, 0, 0, hWnd)
    If (fShowError And rc <> 0) Then
        mciGetErrorString rc, errStr, Len(errStr)
        MsgBox errStr
    End If
    SendMCIString = (rc = 0)
End Function

Private Sub Form_Load()
    Dim rc  As Long
    Dim OK As Boolean

    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If MMSYSERR_NOERROR <> rc Then
        MsgBox "Could not open the mixer.", vbCritical, "Volume Control"
        Exit Sub
    End If

    OK = fGetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)

    If OK Then
        With Volume
            .Max = volCtrl.lMinimum
            .Min = volCtrl.lMaximum \ 2
            .SmallChange = 1000
            .LargeChange = 1000
        End With
    End If
    
    FastForwardSpeed = 10
    CDLoad = False
    
    If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
        
    End If
    
    SendMCIString "set cd time format tmsf wait", True
    MsgBox ("Open CD rom Drive.")
    SendMCIString "set cd door open", True  'sets cd door open
    MsgBox ("Put your compact disk in the CD Rom drive and click Close.")

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close all MCI devices opened by this program
SendMCIString "close all", False
End Sub

Private Sub Image1_Click()
    SendMCIString "play cd", True
    Playing = True
End Sub

Private Sub Image2_Click()
    SendMCIString "stop cd wait", True
    Command = "seek cd to " & Track
    SendMCIString Command, True
    Playing = False
    Update
End Sub

Private Sub Image3_Click()
    SendMCIString "pause cd", True
    Playing = False
    Update
End Sub

Private Sub Image4_Click()
    Dim e As String * 40
    
    SendMCIString "set cd time format milliseconds", True
    mciSendString "status cd position wait", e, Len(e), 0
    If (Playing) Then
        Command = "play cd from " & CStr(CLng(e) - FastForwardSpeed * 1000)
    Else
        Command = "seek cd to " & CStr(CLng(e) - FastForwardSpeed * 1000)
    End If
    mciSendString Command, 0, 0, 0
    SendMCIString "set cd time format tmsf", True
    Update
End Sub

Private Sub Image5_Click()
    Dim e As String * 40
    
    SendMCIString "set cd time format milliseconds", True
    mciSendString "status cd position wait", e, Len(e), 0
    If (Playing) Then
        Command = "play cd from " & CStr(CLng(e) + FastForwardSpeed * 1000)
    Else
        Command = "seek cd to " & CStr(CLng(e) + FastForwardSpeed * 1000)
    End If
    mciSendString Command, 0, 0, 0
    SendMCIString "set cd time format tmsf", True
    Update
End Sub

Private Sub Image6_Click()
    If CDROMopen = False Then
        SendMCIString "set cd door open", True
        CDROMopen = True
        Update
    ElseIf CDROMopen = True Then
        SendMCIString "set cd door closed", True
        CDROMopen = False
        Update
    End If
End Sub

Private Sub Image7_Click()
    Dim from As String
    If (Minute = 0 And Second = 0) Then
        If (Track > 1) Then
            from = CStr(Track - 1)
        Else
            from = CStr(TotalTracks)
        End If
    Else
        from = CStr(Track)
    End If
    If (Playing) Then
        Command = "play cd from " & from
        SendMCIString Command, True
    Else
        Command = "seek cd to " & from
        SendMCIString Command, True
    End If
    Update
End Sub

Private Sub Update()
    Static e As String * 30
    
    mciSendString "status cd media present", e, Len(e), 0
    If (CBool(e)) Then

        If (CDLoad = False) Then
            mciSendString "status cd number of tracks wait", e, Len(e), 0
            TotalTracks = CInt(Mid$(e, 1, 2))
            
            ' If CD only has 1 track, then it's probably a data CD
            If (TotalTracks = 1) Then
                Exit Sub
            End If
            
            mciSendString "status cd length wait", e, Len(e), 0
            TotalTrack.Caption = "Tracks: " & TotalTracks & "  Total time: " & e
            ReDim TrackLength(1 To TotalTracks)
            Dim i As Integer
            For i = 1 To TotalTracks
                Command = "status cd length track " & i
                mciSendString Command, e, Len(e), 0
                TrackLength(i) = e
            Next
    
            CDLoad = True
            SendMCIString "seek cd to 1", True
        End If
    
        ' Update the track time display
        mciSendString "status cd position", e, Len(e), 0
        Track = CInt(Mid$(e, 1, 2))
        Minute = CInt(Mid$(e, 4, 2))
        Second = CInt(Mid$(e, 7, 2))
        TimeWindow.Text = "[" & Format(Track, "00") & "] " & Format(Minute, "00") & ":" & Format(Second, "00")
        TrackTime.Caption = "Track time: " & TrackLength(Track)
    
        ' Check if CD is playing
        mciSendString "status cd mode", e, Len(e), 0
        Playing = (Mid$(e, 1, 7) = "playing")
    Else

        ' Disable all the controls, clear the display
        If (CDLoad = True) Then
            Play.Enabled = False
            Pause.Enabled = False
            FastForward.Enabled = False
            Rewind.Enabled = False
            NextTrack.Enabled = False
            PreviousTrack.Enabled = False
            stpButton.Enabled = False
            CDLoad = False
            Playing = False
            TrackTime.Caption = ""
            TrackTime.Caption = ""
            TimeWindow.Text = ""
        End If
    End If
End Sub

Private Function fSetVolumeControl(ByVal hmixer As Long, _
    mxc As MIXERCONTROL, ByVal Volume As Long) As Boolean
'
' This function sets the value for a volume control.
'
Dim rc   As Long
Dim mxcd As MIXERCONTROLDETAILS
Dim vol  As MIXERCONTROLDETAILS_UNSIGNED

With mxcd
    .item = 0
    .dwControlID = mxc.dwControlID
    .cbStruct = Len(mxcd)
    .cbDetails = Len(vol)
End With
'
' Allocate a buffer for the control value buffer.
'
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
mxcd.cChannels = 1
vol.dwValue = Volume
'
' Copy the data into the control value buffer.
'
Call CopyPtrFromStruct(mxcd.paDetails, vol, Len(vol))
'
' Set the control value.
'
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
Call GlobalFree(hmem)

If MMSYSERR_NOERROR = rc Then
    fSetVolumeControl = True
Else
    fSetVolumeControl = False
End If
End Function

Private Sub Image8_Click()
    If (Track < TotalTracks) Then
        If (Playing) Then
            Command = "play cd from " & Track + 1
            SendMCIString Command, True
        Else
            Command = "seek cd to " & Track + 1
            SendMCIString Command, True
        End If
    Else
        SendMCIString "seek cd to 1", True
    End If
    Update
End Sub

Private Sub Timer1_Timer()
    Update
End Sub

Private Sub Volume_Change()
    lblVolume.Visible = True
    Dim lVol As Long
    lVol = CLng(Volume.Value) * 2
    Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub

Private Sub Volume_Scroll()
    Dim lVol As Long
    
    lVol = CLng(Volume.Value) * 2
    Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub
