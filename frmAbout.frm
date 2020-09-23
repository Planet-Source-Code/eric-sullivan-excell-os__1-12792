VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CD Player portion Copyright (c) Evan Silich."
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
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   3375
      Left            =   120
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   3420
      Left            =   105
      Top             =   105
      Width           =   4740
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3480
      Left            =   75
      Top             =   75
      Width           =   4800
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   3525
      Left            =   45
      Top             =   45
      Width           =   4860
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3600
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label lblcontact 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This code is completely free, you may use whatever parts you wish, without contacting me for consent."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label lblrel 
      BackStyle       =   0  'Transparent
      Caption         =   "Release Date: 2000/10/26"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblver 
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0b"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3615
      Left            =   4920
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   120
      Top             =   3600
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imglogo_Click()
Unload Me

End Sub

Private Sub lblCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me

End Sub

Private Sub picCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Form_Load()
    AutoRedraw = True
    ScaleMode = vbPixels

    With Font
        .Name = "Verdana"
        .Bold = True
        .Size = 40
    End With
    
    ForeColor = &H808080
    CurrentX = 10
    CurrentY = 10
    Print " Excell OS"

    ForeColor = vbBlue
    CurrentX = 10 - 3
    CurrentY = 10 - 3
    Print " Excell OS"
End Sub
