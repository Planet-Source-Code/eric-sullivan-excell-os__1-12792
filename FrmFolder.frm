VERSION 5.00
Begin VB.Form FrmFolder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label1"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3120
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   3045
      Left            =   45
      Top             =   45
      Width           =   4860
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3000
      Left            =   75
      Top             =   75
      Width           =   4800
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   2940
      Left            =   105
      Top             =   105
      Width           =   4740
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   2895
      Left            =   120
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3015
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
      Top             =   3120
      Width           =   4935
   End
End
Attribute VB_Name = "FrmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
