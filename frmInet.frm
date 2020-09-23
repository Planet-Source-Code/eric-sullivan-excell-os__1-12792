VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmInet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   2625
   ClientTop       =   2475
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12990
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
      Left            =   12240
      TabIndex        =   5
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
      Left            =   12480
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   7575
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   12375
      ExtentX         =   21828
      ExtentY         =   13361
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
   Begin VB.ComboBox cmbaddr 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInet.frx":0000
      Left            =   240
      List            =   "frmInet.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "http://www.planet-source-code.com/vb"
      Top             =   600
      Width           =   11775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "ExcelOS WebBrowser v1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   135
      Width           =   2490
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   8655
      Left            =   105
      Top             =   120
      Width           =   12625
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   8700
      Left            =   105
      Top             =   105
      Width           =   12660
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   8760
      Left            =   75
      Top             =   75
      Width           =   12720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   8805
      Left            =   45
      Top             =   45
      Width           =   12780
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   8880
      Left            =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   12615
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   120
      Top             =   8880
      Width           =   12855
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8775
      Left            =   12840
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub drv_Change()
'web.Navigate drv.Drive
'End Sub

Private Sub Command1_Click()
web.Navigate cmbaddr.Text
cmbaddr.AddItem cmbaddr.Text

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub OptBtn_Click(Index As Integer)
    Me.Visible = False
End Sub
