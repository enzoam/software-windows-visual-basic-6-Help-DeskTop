VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmtoolbar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Barra de Ferramentas Help Desk TOP"
   ClientHeight    =   390
   ClientLeft      =   8340
   ClientTop       =   1605
   ClientWidth     =   3180
   LinkTopic       =   "Frmtool"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3180
   Begin ComctlLib.Toolbar tbrtoolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "lsttoolbar"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "indexar"
            Description     =   "Indexar chamado"
            Object.ToolTipText     =   "Reindexar Chamados"
            Object.Tag             =   "btnindexar"
            ImageKey        =   "indexar"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "relogio"
            Description     =   "Relogio do Sistema"
            Object.ToolTipText     =   "Relogio do Sistema"
            Object.Tag             =   "btnrelogio"
            ImageKey        =   "relogio"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "atualizar"
            Object.ToolTipText     =   "Atualizar Chamados"
            Object.Tag             =   "btnautoatualizar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pesquisachamado"
            Object.ToolTipText     =   "Pesquisa de Chamados"
            Object.Tag             =   "pesquisachama"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "trocausuario"
            Object.ToolTipText     =   "Alternar usuário"
            Object.Tag             =   "btnchangeuser"
            ImageIndex      =   4
         EndProperty
      EndProperty
      MouseIcon       =   "tbarhd.frx":0000
   End
   Begin ComctlLib.ImageList lsttoolbar 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tbarhd.frx":001C
            Key             =   "indexar"
            Object.Tag             =   "indexar"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tbarhd.frx":0336
            Key             =   "relogio"
            Object.Tag             =   "relogio"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tbarhd.frx":0650
            Key             =   "copiabanco"
            Object.Tag             =   "copiabanco"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tbarhd.frx":096A
            Key             =   "usuario"
            Object.Tag             =   "usuario"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tbarhd.frx":0C84
            Key             =   "reparar"
            Object.Tag             =   "reparar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmtoolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tbrtoolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    If Button.Tag = "btnindexar" Then
        If frmmain.btnchamaadd.Enabled = True Then
            frmmain.dtamainfile.RecordSource = "Select * from chamados Order by chama_num"
            frmmain.dtamainfile.Refresh
        Else
            MsgBox "O arquivo de Chamados está aberto", vcokonly, "Não foi possível reindexar os registros"
        End If
    End If
    If Button.Tag = "btnrelogio" Then
        MsgBox "A hora atual em seu sistema é " & Time, vbOKOnly, "Informação"
    End If
        If Button.Tag = "pesquisachama" Then
            If frmmain.mnu_pesquisa_chama.Item(1).Checked = True Then
            frmmain.frmpesquisachama.Enabled = False
            frmmain.frmpesquisachama.Visible = False
            frmmain.mnu_pesquisa_chama.Item(1).Checked = False
        Else
            frmmain.frmpesquisachama.Enabled = True
            frmmain.frmpesquisachama.Visible = True
            frmmain.mnu_pesquisa_chama.Item(1).Checked = True
        End If
    End If
    If Button.Tag = "btnautoatualizar" Then
        frmmain.dtamainfile.Refresh
    End If
    If Button.Tag = "btnchangeuser" Then
        formentrasenha.Show
        Unload frmmain
    End If
End Sub
