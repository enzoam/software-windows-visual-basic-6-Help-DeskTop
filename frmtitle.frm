VERSION 5.00
Begin VB.Form frmtitle 
   BorderStyle     =   0  'None
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Projeto - Tecnologia em Processamento de Dados -  Programação de Computadores II "
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Beta v 1.5  -  Novembro de 2000"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   840
      Picture         =   "frmtitle.frx":0000
      Top             =   0
      Width           =   5280
   End
   Begin VB.Image Image2 
      Height          =   2130
      Left            =   1560
      Picture         =   "frmtitle.frx":36DA
      Top             =   1080
      Width           =   3990
   End
End
Attribute VB_Name = "frmtitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Interval = 5000
End Sub

Private Sub Timer1_Timer()
    formentrasenha.Show
    Unload frmtitle
End Sub
