Attribute VB_Name = "modMain"
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub EndApp()
    Unload frmMain
    Unload frmAbout
    Unload frmInet
    End
End Sub

'Public Sub ChangeColours(LabelToChange As Label, OtherLabel1 As Label, OtherLabel2 As Label, OtherLabel3 As Label, OtherLabel4 As Label)
    'LabelToChange.BackColor = &HFFC0C0
    'OtherLabel1.BackColor = &HE0E0E0
    'OtherLabel2.BackColor = &HE0E0E0
    'OtherLabel3.BackColor = &HE0E0E0
    'OtherLabel4.BackColor = &HE0E0E0
'End Sub

'Public Sub ResetColours(LabelToChange As Label, OtherLabel1 As Label, OtherLabel2 As Label, OtherLabel3 As Label, OtherLabel4 As Label)
    'LabelToChange.BackColor = &HE0E0E0
    'OtherLabel1.BackColor = &HE0E0E0
    'OtherLabel2.BackColor = &HE0E0E0
    'OtherLabel3.BackColor = &HE0E0E0
    'OtherLabel4.BackColor = &HE0E0E0
'End Sub

