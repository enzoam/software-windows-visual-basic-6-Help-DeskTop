VERSION 5.00
Begin VB.Form frmbdusers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fun��es / Usu�rios"
   ClientHeight    =   4200
   ClientLeft      =   2880
   ClientTop       =   1650
   ClientWidth     =   6495
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dtaeq 
      Connect         =   "Access"
      DatabaseName    =   "D:\Arquivo\Faculdade\Programa��o de Computadores\3� ano\Programa\Cadastro\hpdsk.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Equipe"
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtuseradd 
      DataField       =   "username"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton btnuseralt 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btnuseradd 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton btnusersave 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton btnuserdel 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtuserpass 
      DataField       =   "userpass"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.CheckBox chkuser 
      Caption         =   "Usu�rio"
      DataField       =   "user"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CheckBox chkadmin 
      Caption         =   "Administrador"
      DataField       =   "admin"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CheckBox chkuserprobs 
      Caption         =   "Acesso a alimenta��o de problemas e solu��es"
      DataField       =   "userprobs"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CheckBox chkuseraltcads 
      Caption         =   "Acesso as fun��es das guias de cadastros"
      DataField       =   "useraltcads"
      DataSource      =   "dtausers"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Data dtausers 
      Connect         =   "Access"
      DatabaseName    =   "D:\Arquivo\Faculdade\Programa��o de Computadores\3� ano\Programa\Cadastro\profiles.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "profile"
      Top             =   3720
      Width           =   3735
   End
   Begin VB.CommandButton btnusercancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton btnokbdusers 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Senha :"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblfuncoestitle 
      Caption         =   "Fun��es para os bancos de cadastro :"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label lblusername 
      Caption         =   "Nome do Usu�rio :"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmbdusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnokbdusers_Click()
    On Error GoTo trataerros:
    If btnuseradd.Enabled = False Then
        MsgBox "O usu�rio est� em aberto", vbCritical, "Aviso ao Administrador"
    Else
        Unload frmbdusers
    End If
trataerros:
    Select Case Err.Number
        Case 3058
            dtausers.Recordset.CancelUpdate
    End Select
End Sub

Private Sub Form_Load()
    dtausers.DatabaseName = "profiles.mdb"
    dtausers.RecordSource = "profile"
    dtaeq.DatabaseName = "hpdsk.mdb"
    dtaeq.RecordSource = "equipe"
End Sub
    
'----------------------------------------------------------------
'Comandos Referentes ao cadastro, bot�es e caixas de controle do
'banco de dados (Senhas e usu�rios)
'----------------------------------------------------------------

Private Sub btnuseradd_Click()
' On Error GoTo trataerros
    dtausers.Recordset.AddNew
    display_user_controls
    btnusersave.Enabled = True
    btnusercancel.Enabled = True
    btnuseradd.Enabled = False
    btnuserdel.Enabled = False
    btnuseralt.Enabled = False
trataerros:
    Select Case Err.Number
        Case 3426
    End Select
End Sub

Private Sub btnuseralt_Click() 'Bot�o - editar registro de clientes
    On Error GoTo trata_erros
    dtausers.Recordset.Edit
    display_user_controls
    btnusersave.Enabled = True
    btnusercancel.Enabled = True
    btnuseradd.Enabled = False
    btnuserdel.Enabled = False
    btnuseralt.Enabled = False
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro n�o encontrado"
            hide_user_controls
         Case 3260
            MsgBox "Este arquivo est� sendo usado por outro usu�rio", vbOKOnly, "Problemas de rede"
            hide_user_controls
    End Select
End Sub

Private Sub btnusercancel_Click()
    On Error GoTo trataerros
    dtausers.Recordset.CancelUpdate
    btnuseradd.Enabled = True
    btnuserdel.Enabled = True
    btnuseralt.Enabled = True
    btnusersave.Enabled = False
    btnusercancel.Enabled = False
    hide_user_controls
    chkadmin.Value = 0
    chkuser.Value = 0
    dtausers.Refresh
trataerros:
    Select Case Err.Number
        Case 444
    End Select
End Sub

Private Sub btnuserdel_Click() 'Bot�o - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtausers.Recordset.Delete
        chkadmin.Value = 0
        chkuser.Value = 0
        dtausers.Refresh
    End If
trata_erros:
    Select Case Err.Number
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro n�o encontrado"
    End Select
End Sub

Private Sub btnusersave_Click() 'Bot�o - salvar registro ou atualizar modifica��es no registro de clientes
    On Error GoTo trata_erros
    dtausers.Recordset.Update
    dtausers.Refresh
    btnuseradd.Enabled = True
    btnuserdel.Enabled = True
    btnuseralt.Enabled = True
    btnusersave.Enabled = False
    btnusercancel.Enabled = False
    hide_user_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "N�o � poss�vel gravar um registro sem um c�digo", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse c�digo j� est� cadastrado", vbOKOnly, "N�o pode gravar o registro"
            btnuseradd.Enabled = True
            btnuserdel.Enabled = True
            btnuseralt.Enabled = True
            btnusersave.Enabled = False
            btnusercancel.Enabled = False
            hide_user_controls
            dtausers.Recordset.MoveFirst
        Case 3260
            MsgBox "Um outro usu�rio est� gravando dados neste registro. Consulte o hist�rico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub
Private Sub display_user_controls() 'Fun��o para habilitar as caixas de texto dos registros de clientes
    txtuseradd.Enabled = True
    txtuserpass.Enabled = True
    chkadmin.Enabled = True
    chkuser.Enabled = True
    chkuseraltcads.Enabled = True
    chkuserprobs.Enabled = True
    dtausers.Enabled = False
End Sub

Private Sub hide_user_controls() 'Fun��o para desabilitar as caixas de texto dos registros de clientes
    txtuseradd.Enabled = False
    txtuserpass.Enabled = False
    chkadmin.Enabled = False
    chkuser.Enabled = False
    chkuseraltcads.Enabled = False
    chkuserprobs.Enabled = False
    dtausers.Enabled = True
End Sub
