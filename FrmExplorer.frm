VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form FrmExplorer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8865
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OptBtn 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Index           =   1
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton OptBtn 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Index           =   0
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  ExcellOS explorer"
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
      TabIndex        =   2
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5280
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   5205
      Left            =   45
      Top             =   45
      Width           =   7140
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5160
      Left            =   75
      Top             =   75
      Width           =   7080
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   5100
      Left            =   105
      Top             =   105
      Width           =   7020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   5175
      Left            =   7200
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   120
      Top             =   5280
      Width           =   7215
   End
End
Attribute VB_Name = "FrmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    drv.Drive = "C:\"
    web.Navigate2 "C:\"
    web.Visible = True
    drv.Visible = True
End Sub

Private Sub drv_Change()
    web.Navigate2 (drv.Drive)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub

Private Sub OptBtn_Click(Index As Integer)
    Me.Visible = False
End Sub
