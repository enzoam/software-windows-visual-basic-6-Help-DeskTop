VERSION 5.00
Begin VB.Form formentrasenha 
   BorderStyle     =   0  'None
   Caption         =   "Help Desk TOP "
   ClientHeight    =   2490
   ClientLeft      =   3060
   ClientTop       =   2190
   ClientWidth     =   5985
   ScaleHeight     =   2490
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtmostranome 
      DataField       =   "username"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CheckBox chkuser 
      DataField       =   "user"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox chkadmin 
      DataField       =   "admin"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox chkprobs 
      DataField       =   "userprobs"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox chkaltcads 
      DataField       =   "useraltcads"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtmostrasenha 
      DataField       =   "userpass"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.Data dtausers 
      Connect         =   "Access"
      DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\profiles.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "profile"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtnome 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton btnlimpar 
         Caption         =   "Limpar"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton btncancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton btnentrar 
         Caption         =   "Entrar"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtsenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
      Begin VB.PictureBox picchave 
         Height          =   495
         Left            =   240
         Picture         =   "formsenha.frx":0000
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblsenha 
         Caption         =   "SENHA :"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbllogin 
         Caption         =   "LOGIN :"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblexplica 
         Caption         =   "Entre com o nome do usuário e a Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "formentrasenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btncancelar_Click()
    Unload formentrasenha
End Sub



Private Sub btnentrar_Click()
    If txtnome = txtmostranome Then
        If txtsenha.Text = txtmostrasenha.Text Then
            frmmain.lblchkadmin.Caption = chkadmin.Value
            frmmain.lblchkuser.Caption = chkuser.Value
            frmmain.lbluseraltcads.Caption = chkaltcads.Value
            frmmain.lbluserprobs.Caption = chkprobs.Value
            frmmain.lblusername = txtnome.Text
            If frmmain.lblchkadmin = 0 Then
               frmmain.menu_admin.Item(1).Enabled = False
            End If
            If frmmain.lblchkuser = 1 Then
               frmmain.btnchamadel.Enabled = False
               frmmain.menu_chama_del.Item(1).Enabled = False
            End If
            If frmmain.lbluseraltcads = 0 Then
                frmmain.btntabcargoadd.Enabled = False
                frmmain.btntabcliadd.Enabled = False
                frmmain.btntabeqadd.Enabled = False
                frmmain.btntabniveladd.Enabled = False
                frmmain.btntabprobadd.Enabled = False
                frmmain.btntabsoladd.Enabled = False
                frmmain.btntabcargodel.Enabled = False
                frmmain.btntabclidel.Enabled = False
                frmmain.btntabeqdel.Enabled = False
                frmmain.btntabniveldel.Enabled = False
                frmmain.btntabprobdel.Enabled = False
                frmmain.btntabsoldel.Enabled = False
                frmmain.btntabcargoalt.Enabled = False
                frmmain.btntabclialt.Enabled = False
                frmmain.btntabeqalt.Enabled = False
                frmmain.btntabnivelalt.Enabled = False
                frmmain.btntabprobalt.Enabled = False
                frmmain.btntabsolalt.Enabled = False
            End If
            frmmain.Show
            Unload formentrasenha
        Else
            MsgBox "Senha incorreta", vcokonly, "Aviso de segurança"
        End If
    Else
        MsgBox "Usuário não autorizado", vcokonly, "Aviso de segurança"
    End If
End Sub

Private Sub btnlimpar_Click()
    txtnome.Text = ""
    txtsenha.Text = ""
End Sub

Private Sub Form_Load()
    dtausers.DatabaseName = "profiles.mdb"
    dtausers.RecordSource = "profile"
End Sub

Private Sub txtnome_Change()
On Error GoTo trataerros
    dtausers.Recordset.FindFirst "username = '" + txtnome.Text + "'"
trataerros:
    Select Case Err.Number
    End Select
End Sub

