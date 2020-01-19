VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmrestore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Restauração dos arquivos de dados"
   ClientHeight    =   2010
   ClientLeft      =   3780
   ClientTop       =   3810
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkprofiles 
      Caption         =   "Lista de usuários"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CheckBox chkhist 
      Caption         =   "Arquivo de registros do Historico"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CheckBox chkchama 
      Caption         =   "Cadastros e Chamados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton btnrestini 
      Caption         =   "Iniciar Restauração"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin ComctlLib.ProgressBar prbrest 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "frmrestore.frx":0000
   End
End
Attribute VB_Name = "frmrestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnrestini_Click()
    On Error GoTo trataerros
    btnrestini.Enabled = False
    btnrestini.Caption = "Efetuando Restauração"
    If chkchama.Value = 1 Then
        FileCopy "hpds_bck.mdb", "hpdsk.mdb"
    End If
    prbrest.Value = 30
    If chkhist.Value = 1 Then
        FileCopy "hist_bck.mdb", "hist.mdb"
    End If
    prbrest.Value = 60
    If chkprofiles.Value = 1 Then
        FileCopy "prof_bck.mdb", "profiles.mdb"
    End If
    prbrest.Value = 100
    Unload frmrestore
trataerros:
    Select Case Err.Number
        Case 70
            MsgBox "O arquivo está aberto", vbCritical, "Aviso do sistema"
            Unload frmrestore
        Case 53
            MsgBox "Arquivo de reposição inexistente", vbCritical, "Aviso do sistema"
            Unload frmrestore
    End Select
End Sub

Private Sub Form_Load()
    Unload frmmain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    formentrasenha.Show
End Sub

