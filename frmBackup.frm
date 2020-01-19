VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copia de segurança dos arquivos de dados"
   ClientHeight    =   960
   ClientLeft      =   3780
   ClientTop       =   3420
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnbckini 
      Caption         =   "Iniciar Backup"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin ComctlLib.ProgressBar prbbackup 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327680
      Appearance      =   1
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnbckini_Click()
    On Error GoTo trataerros
    btnbckini.Enabled = False
    btnbckini.Caption = "Efetuando Backup"
    FileCopy "hpdsk.mdb", "hpds_bck.mdb"
    prbbackup.Value = 30
    FileCopy "hist.mdb", "hist_bck.mdb"
    prbbackup.Value = 60
    FileCopy "profiles.mdb", "prof_bck.mdb"
    prbbackup.Value = 100
    Unload frmBackup
trataerros:
    Select Case Err.Number
        Case 70
            MsgBox "O arquivo está aberto", vbCritical, "Aviso do sistema"
    End Select
End Sub

Private Sub Form_Load()
    Unload frmmain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    formentrasenha.Show
End Sub
