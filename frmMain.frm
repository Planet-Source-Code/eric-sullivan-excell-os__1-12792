VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MacShell 1.0.0"
   ClientHeight    =   12960
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox ChangeLabelCap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   12480
      Width           =   17055
      Begin VB.OptionButton Option1 
         Caption         =   "Desktop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
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
      Left            =   360
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Label LblFolder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Folder #1"
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
      Index           =   0
      Left            =   6840
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image ImgFolder 
      Height          =   480
      Index           =   0
      Left            =   7200
      Picture         =   "frmMain.frx":08CA
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Folder"
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
      Left            =   6960
      TabIndex        =   12
      Top             =   470
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   12735
      Left            =   120
      Top             =   120
      Width           =   17055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rename"
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
      Index           =   2
      Left            =   3240
      TabIndex        =   10
      Top             =   470
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Move..."
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
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   470
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   470
      Width           =   615
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   5
      Left            =   360
      Picture         =   "frmMain.frx":1194
      Top             =   6240
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Games"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   6720
      Width           =   720
   End
   Begin VB.Label lblDeskTrash 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trash Can"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Image ImgTrash 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmMain.frx":1A5E
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image ImgTrash 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":2328
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Internet  Browser"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   4
      Left            =   360
      Picture         =   "frmMain.frx":2BF2
      ToolTipText     =   "Internet Browser - Use this to access the world wide web!"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Media  Player"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   3
      Left            =   360
      Picture         =   "frmMain.frx":34BC
      ToolTipText     =   "Media Player - Listen to your favourite audio files."
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   450
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":3D86
      ToolTipText     =   "Run - Select a file or appilcation and run it."
      Top             =   600
      Width           =   480
   End
   Begin VB.Label DesktopLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Explorer  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmMain.frx":4650
      ToolTipText     =   "My Mac - Explore your computer!"
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Time and Date will go here...!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11280
      TabIndex        =   6
      Top             =   120
      Width           =   5820
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   12795
      Left            =   90
      Top             =   90
      Width           =   17115
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   12850
      Left            =   60
      Top             =   60
      Width           =   17175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   12915
      Left            =   30
      Top             =   30
      Width           =   17235
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   12975
      Left            =   0
      Top             =   0
      Width           =   17295
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   17055
   End
   Begin VB.Label DesktopLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   390
      TabIndex        =   2
      Top             =   7680
      Width           =   435
   End
   Begin VB.Image DesktopImage 
      Height          =   480
      Index           =   2
      Left            =   360
      Picture         =   "frmMain.frx":4F1A
      ToolTipText     =   "Find - Find what you wan't, fast and easily."
      Top             =   7200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   13440
      Picture         =   "frmMain.frx":57E4
      ToolTipText     =   "Click here for options!"
      Top             =   600
      Width           =   3525
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   855
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   855
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   2
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   975
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   855
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Long, OldY As Long, IsMoving As Boolean
Dim Selected As Integer, Stuffin As Boolean
Dim ChangingCaption As Boolean

Private Sub ChangeLabelCap_Change()
    ChangingCaption = True
    ChrNum = Len(ChangeLabelCap)
    Select Case ChrNum
        Case 13: ChangeLabelCap.Height = 525: Label1.Height = 525
        Case 26: ChangeLabelCap.Height = 765: Label1.Height = 765
        Case 39: ChangeLabelCap.Height = 1005: Label1.Height = 1005
    End Select
End Sub

Private Sub DesktopImage_Click(Index As Integer)
    Select Case Index
        Case 0: Call ChangeStyle(0)
        Case 1: Call ChangeStyle(1)
        Case 2: Call ChangeStyle(2)
        Case 3: Call ChangeStyle(3)
        Case 4: Call ChangeStyle(4)
        Case 5: Call ChangeStyle(5)
    End Select
    
    Selected = DesktopLabel(Index).Index
End Sub

Private Sub DesktopImage_DblClick(Index As Integer)
    Select Case Index
        Case 0: FrmExplorer.Visible = True
        Case 1: 'Call ShowRunDialog(Me, "MacSHELL", "Select the file you want to open.")
        Case 2: 'Call ShowFindDialog
        Case 3: frmMed.Visible = True
        Case 4: frmInet.Visible = True
        Case 5
    End Select
    
    Static i As Integer
    i = i + 1
    Load Option1(i)
    
    Option1(i).Left = Option1(i - 1).Left + 1500
    Option1(i).Top = Option1(i - 1).Top
    Option1(i).Caption = DesktopLabel(Selected)
    Option1(i).Visible = True
    
    If Len(Option1(i).Caption) > 13 Then
        makesmaller = Option1(i).Caption
        For Q = 1 To Len(makesmaller)
            Q = Q + 1
            Option1(i).Caption = Mid(makesmaller, Q)
        Next Q
    End If
End Sub

Private Sub DesktopImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: OldX = X: OldY = Y: IsMoving = True
        Case 1: OldX = X: OldY = Y: IsMoving = True
        Case 2: OldX = X: OldY = Y: IsMoving = True
        Case 3: OldX = X: OldY = Y: IsMoving = True
        Case 4: OldX = X: OldY = Y: IsMoving = True
        Case 5: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub DesktopImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            If IsMoving Then
                DesktopImage(Num).Top = DesktopImage(Num).Top - (OldY - Y)
                DesktopImage(Num).Left = DesktopImage(Num).Left - (OldX - X)
        
                DesktopLabel(Num).Top = DesktopLabel(Num).Top - (OldY - Y)
                DesktopLabel(Num).Left = DesktopLabel(Num).Left - (OldX - X)
            End If
            
        Case 1
            If IsMoving Then
                DesktopImage(1).Top = DesktopImage(1).Top - (OldY - Y)
                DesktopImage(1).Left = DesktopImage(1).Left - (OldX - X)
        
                DesktopLabel(1).Top = DesktopLabel(1).Top - (OldY - Y)
                DesktopLabel(1).Left = DesktopLabel(1).Left - (OldX - X)
            End If
        
        Case 2
            If IsMoving Then
                DesktopImage(2).Top = DesktopImage(2).Top - (OldY - Y)
                DesktopImage(2).Left = DesktopImage(2).Left - (OldX - X)
        
                DesktopLabel(2).Top = DesktopLabel(2).Top - (OldY - Y)
                DesktopLabel(2).Left = DesktopLabel(2).Left - (OldX - X)
            End If
        
        Case 3
            If IsMoving Then
                DesktopImage(3).Top = DesktopImage(3).Top - (OldY - Y)
                DesktopImage(3).Left = DesktopImage(3).Left - (OldX - X)
        
                DesktopLabel(3).Top = DesktopLabel(3).Top - (OldY - Y)
                DesktopLabel(3).Left = DesktopLabel(3).Left - (OldX - X)
            End If
            
        Case 4
            If IsMoving Then
                DesktopImage(4).Top = DesktopImage(4).Top - (OldY - Y)
                DesktopImage(4).Left = DesktopImage(4).Left - (OldX - X)
        
                DesktopLabel(4).Top = DesktopLabel(4).Top - (OldY - Y)
                DesktopLabel(4).Left = DesktopLabel(4).Left - (OldX - X)
            End If
        
        Case 5
            If IsMoving Then
                DesktopImage(5).Top = DesktopImage(5).Top - (OldY - Y)
                DesktopImage(5).Left = DesktopImage(5).Left - (OldX - X)
        
                DesktopLabel(5).Top = DesktopLabel(5).Top - (OldY - Y)
                DesktopLabel(5).Left = DesktopLabel(5).Left - (OldX - X)
            End If
        
    End Select
End Sub

Private Sub DesktopImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: IsMoving = False
        Case 1: IsMoving = False
        Case 2: IsMoving = False
        Case 3: IsMoving = False
        Case 4: IsMoving = False
        Case 5: IsMoving = False
    End Select
End Sub

Private Sub Form_Click()
    For Num = 0 To 5
        DesktopLabel(Num).FontBold = False
    Next Num
    
    For i = 0 To 2
        Label3(i).Visible = False
        Shape9(i).Visible = False
    Next i
    
    If ChangingCaption = True Then
        DesktopLabel(Selected).Caption = ChangeLabelCap.Text
        ChangeLabelCap.Visible = False
        DesktopLabel(Selected).Visible = True
    Else
    End If
End Sub

Private Sub Form_Load()
    Label1.Caption = Date & " " & Time
    
    For i = 0 To 2
        Label3(i).Visible = False
        Shape9(i).Visible = False
    Next i
End Sub

Private Sub Image1_Click()
    frmAbout.Visible = True
End Sub

Private Sub ImgFolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub ImgFolder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected
            If IsMoving Then
                ImgFolder(Selected).Top = ImgFolder(Selected).Top - (OldY - Y)
                ImgFolder(Selected).Left = ImgFolder(Selected).Left - (OldX - X)
        
                LblFolder(Selected).Top = LblFolder(Selected).Top - (OldY - Y)
                LblFolder(Selected).Left = LblFolder(Selected).Left - (OldX - X)
            End If
    End Select
End Sub

Private Sub ImgFolder_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected: IsMoving = False
    End Select
End Sub

Private Sub imgsd_Click()
    End
End Sub

Private Sub imgrun_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then

    ElseIf Button = 1 Then
        OldX = X
        OldY = Y
        IsMoving = True
    End If
End Sub

Private Sub imgrun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsMoving Then
        imgrun.Top = imgrun.Top - (OldY - Y)
        imgrun.Left = imgrun.Left - (OldX - X)
        
        lbldeskrun.Top = lbldeskrun.Top - (OldY - Y)
        lbldeskrun.Left = lbldeskrun.Left - (OldX - X)
    End If
End Sub

Private Sub imgrun_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsMoving = False
End Sub

Private Sub ImgTrash_DblClick(Index As Integer)
    If Stuffin = True Then
        TrashContents.Visible = True
    End If
End Sub

Private Sub Label2_Click()
EndApp
End Sub

Private Sub Label3_Click(Index As Integer)
    If Label3(Index).Index = 0 Then
        MsgVar = MsgBox("Are you sure you want to delete " & Chr(34) & DesktopLabel(Selected).Caption & Chr(34) & " to the trash?", vbYesNo + vbQuestion, "Delete Confermation")
        Select Case MsgVar
            Case vbYes
                DesktopLabel(Selected).Visible = False
                DesktopImage(Selected).Visible = False
                If ImgTrash(0).Visible = True Then
                    ImgTrash(0).Visible = False
                    ImgTrash(1).Visible = True
                End If
                Stuffin = True
            Case vbNo
                Cancel = Not ReadyToDelete
        End Select
    ElseIf Label3(Index).Index = 1 Then
    ElseIf Label3(Index).Index = 2 Then
        ChangeLabelCap.Left = DesktopLabel(Selected).Left
        ChangeLabelCap.Top = DesktopLabel(Selected).Top
        ChangeLabelCap.Visible = True
        DesktopLabel(Selected).Visible = False
        ChangeLabelCap.SetFocus
    End If
End Sub

Private Sub Label4_Click()
    Static i As Integer
    i = i + 1
    Load ImgFolder(i)
    Load LblFolder(i)
    
    ImgFolder(i).Left = ImgFolder(i - 1).Left + 200
    ImgFolder(i).Top = ImgFolder(i - 1).Top + 600

    LblFolder(i).Left = LblFolder(i - 1).Left + 200
    LblFolder(i).Top = LblFolder(i - 1).Top + 600

    LblFolder(i).Caption = "New Folder #" & i
    ImgFolder(i).Visible = True
    LblFolder(i).Visible = True
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Date & " " & Time
End Sub

Private Sub ChangeStyle(Num As Integer)
    For i = 0 To 5
        DesktopLabel(i).BorderStyle = 0
        DesktopLabel(i).FontBold = False
    Next i
            
    DesktopLabel(Num).FontBold = True

    For i = 0 To 2
        Label3(i).Visible = True
        Shape9(i).Visible = True
    Next i
End Sub

'Public Sub PlaySound(strFileName As String)
'    sndPlaySound strFileName, 1
'End Sub
