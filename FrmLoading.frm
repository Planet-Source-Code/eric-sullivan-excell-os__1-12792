VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLoading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   1680
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   2055
      Left            =   120
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   2100
      Left            =   105
      Top             =   105
      Width           =   4620
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2160
      Left            =   75
      Top             =   75
      Width           =   4680
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   2205
      Left            =   45
      Top             =   45
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   2275
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   240
      Picture         =   "FrmLoading.frx":0000
      Top             =   240
      Width           =   4350
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    PB1 = PB1 + 1
    If PB1.Value = 30 Then
        Timer1.Interval = "20"
    ElseIf PB1.Value = 100 Then
        Timer1.Enabled = False
        Me.Visible = False
        
        Timer1.Enabled = True
        Timer1.Interval = "2000"
        frmMain.Visible = True
        Timer1.Enabled = False
    End If
End Sub
