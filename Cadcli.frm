VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sistema Integrado - Help Desk TOP"
   ClientHeight    =   8040
   ClientLeft      =   -135
   ClientTop       =   630
   ClientWidth     =   11880
   FillColor       =   &H00E0E0E0&
   Icon            =   "Cadcli.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame frmfundo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordens e Serviços"
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton btnchamaprint 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Imprimir ordem de serviço"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtcodchama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_num"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   20
         ToolTipText     =   "Codigo do chamado - Máximo 6 números"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtdataechama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_data_abre"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Data de inclusão do chamado - xx/xx/xx"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtdataschama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_data_fecha"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   18
         ToolTipText     =   "Data de fechamento do chamado - xx/xx/xx"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtcodcli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_cli"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   17
         ToolTipText     =   "Codigo do Cliente cadastrado - Máximo 6 números"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtnomecli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Nome do Cliente cadastrado"
         Top             =   840
         Width           =   4695
      End
      Begin VB.CommandButton btnchamaadd 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Incluir"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Incluir novo chamado"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton btnchamaalt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Editar"
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Editar chamado atual"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton btnchamasave 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salvar chamado em aberto"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton btnchamadel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excluir"
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir chamado atual"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txtsolchama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_soldesc"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   1125
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         ToolTipText     =   "Descrição da Solução"
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox txthoraschama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_qtd_horas"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         MaxLength       =   9
         TabIndex        =   10
         ToolTipText     =   "Quantidade total de horas utilizadas"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtdescchama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "chama_desc"
         DataSource      =   "dtamainfile"
         Enabled         =   0   'False
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Descrição do chamado"
         Top             =   1680
         Width           =   5055
      End
      Begin VB.CommandButton btnchamacancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancelar alterações"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Frame frmpesquisachama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pesquisa :"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   5640
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox txtchamapesquisa 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Conteúdo de pesquisa"
            Top             =   840
            Width           =   2895
         End
         Begin VB.ComboBox cmbchamapesquisa 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Cadcli.frx":0442
            Left            =   120
            List            =   "Cadcli.frx":045B
            TabIndex        =   2
            ToolTipText     =   "Campos de pesquisa"
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Timer tmrcontrol 
         Left            =   11280
         Top             =   2400
      End
      Begin VB.Data dtamainfile 
         Appearance      =   0  'Flat
         Connect         =   "Access"
         DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\hpdsk.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   9000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Chamados"
         ToolTipText     =   "Chamados"
         Top             =   3720
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog cmndialog 
         Left            =   9000
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDBCtls.DBCombo dbcequipesp 
         Bindings        =   "Cadcli.frx":04D1
         DataField       =   "chama_sp"
         DataSource      =   "dtamainfile"
         Height          =   315
         Left            =   6960
         TabIndex        =   4
         ToolTipText     =   "Componentes da equipe cadastrados"
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   -2147483641
         ListField       =   "eq_nome"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcequiperp 
         Bindings        =   "Cadcli.frx":04E8
         DataField       =   "chama_repas"
         DataSource      =   "dtamainfile"
         Height          =   315
         Left            =   6960
         TabIndex        =   5
         ToolTipText     =   "Componentes da equipe cadastrados"
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   -2147483641
         ListField       =   "eq_nome"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcequipeatp 
         Bindings        =   "Cadcli.frx":04FF
         DataField       =   "chama_atp"
         DataSource      =   "dtamainfile"
         Height          =   315
         Left            =   6960
         TabIndex        =   6
         ToolTipText     =   "Componentes da equipe cadastrados"
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   -2147483641
         ListField       =   "eq_nome"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbctipoprob 
         Bindings        =   "Cadcli.frx":0516
         DataField       =   "chama_tipo_prob"
         DataSource      =   "dtamainfile"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Tipos de problemas disponíveis no Banco Solução"
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "prob_tipo"
         BoundColumn     =   "prob_cod"
         Text            =   ""
      End
      Begin TabDlg.SSTab tabmain 
         Height          =   3015
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   5318
         _Version        =   393216
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Histórico"
         TabPicture(0)   =   "Cadcli.frx":0531
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblresp"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblhstsolsn"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblhsthora"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblhsteq"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblhstnomecli"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblhstprob"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblhstdatas"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblhstdatae"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblhstcodchama"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "dbghst"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "btnhstfiltro"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmbhstfiltro"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txthstn"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txthstsolsn"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txthstresp"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txthsthora"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txthstnomeeq"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txthstnomecli"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txthstdataschama"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txthstdataechama"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txthsttipoprob"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txthstchamacod"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "dtahistorico"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).ControlCount=   23
         TabCaption(1)   =   "Clientes"
         TabPicture(1)   =   "Cadcli.frx":054D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmbtabpesquisacli"
         Tab(1).Control(1)=   "btntabclicancel"
         Tab(1).Control(2)=   "txttabnivcli"
         Tab(1).Control(3)=   "frmpesquisacli"
         Tab(1).Control(4)=   "btntabclisave"
         Tab(1).Control(5)=   "dtacliente"
         Tab(1).Control(6)=   "btntabclialt"
         Tab(1).Control(7)=   "btntabclidel"
         Tab(1).Control(8)=   "btntabcliadd"
         Tab(1).Control(9)=   "txttabcodcli"
         Tab(1).Control(10)=   "txttabnomecli"
         Tab(1).Control(11)=   "txttabendcli"
         Tab(1).Control(12)=   "txttabfonecli"
         Tab(1).Control(13)=   "txttabmailcli"
         Tab(1).Control(14)=   "lbltabnivcli"
         Tab(1).Control(15)=   "lbltanivclii"
         Tab(1).Control(16)=   "lbltabcodcli"
         Tab(1).Control(17)=   "lbetabnomecli"
         Tab(1).Control(18)=   "lbltabendcli"
         Tab(1).Control(19)=   "lbltabfonecli"
         Tab(1).Control(20)=   "lbltabmailcli"
         Tab(1).ControlCount=   21
         TabCaption(2)   =   "Equipe"
         TabPicture(2)   =   "Cadcli.frx":0569
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmbtabpesquisaeq"
         Tab(2).Control(1)=   "btntabeqcancel"
         Tab(2).Control(2)=   "frmtabeqpesquisa"
         Tab(2).Control(3)=   "btntabeqalt"
         Tab(2).Control(4)=   "btntabeqsave"
         Tab(2).Control(5)=   "btntabeqdel"
         Tab(2).Control(6)=   "btntabeqadd"
         Tab(2).Control(7)=   "dtatabeq"
         Tab(2).Control(8)=   "txttabcargoequipe"
         Tab(2).Control(9)=   "txttabmailequipe"
         Tab(2).Control(10)=   "txttabfonequipe"
         Tab(2).Control(11)=   "txttabendequipe"
         Tab(2).Control(12)=   "txttabnomequipe"
         Tab(2).Control(13)=   "txttabcodequipe"
         Tab(2).Control(14)=   "lbltabcargoequipe"
         Tab(2).Control(15)=   "lblcargoequipe"
         Tab(2).Control(16)=   "lblmailquipe"
         Tab(2).Control(17)=   "lblfoneequipe"
         Tab(2).Control(18)=   "lblendequipe"
         Tab(2).Control(19)=   "lblnomequipe"
         Tab(2).Control(20)=   "lblcodequipe"
         Tab(2).ControlCount=   21
         TabCaption(3)   =   "Banco Solução"
         TabPicture(3)   =   "Cadcli.frx":0585
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "btntabsolcopiar"
         Tab(3).Control(1)=   "txttabsolprob"
         Tab(3).Control(2)=   "txttabsolcod"
         Tab(3).Control(3)=   "txttabprobdesc"
         Tab(3).Control(4)=   "txttabprobcod"
         Tab(3).Control(5)=   "btntabsolcancel"
         Tab(3).Control(6)=   "btntabprobcancel"
         Tab(3).Control(7)=   "btntabsoldel"
         Tab(3).Control(8)=   "btntabprobsave"
         Tab(3).Control(9)=   "btntabsolalt"
         Tab(3).Control(10)=   "btntabsolsave"
         Tab(3).Control(11)=   "btntabsoladd"
         Tab(3).Control(12)=   "btntabprobdel"
         Tab(3).Control(13)=   "btntabprobalt"
         Tab(3).Control(14)=   "btntabprobadd"
         Tab(3).Control(15)=   "txttabsoldesc"
         Tab(3).Control(16)=   "dtasoluções"
         Tab(3).Control(17)=   "txttabsolcont"
         Tab(3).Control(18)=   "dtaproblemas"
         Tab(3).Control(19)=   "Label4"
         Tab(3).Control(20)=   "Label3"
         Tab(3).Control(21)=   "Label2"
         Tab(3).Control(22)=   "Label1"
         Tab(3).Control(23)=   "lblcodproblema"
         Tab(3).ControlCount=   24
         TabCaption(4)   =   "Cargos"
         TabPicture(4)   =   "Cadcli.frx":05A1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lbltabcodcargo"
         Tab(4).Control(1)=   "lbltabdesccargo"
         Tab(4).Control(2)=   "txttabdesccargo"
         Tab(4).Control(3)=   "txttabcodcargo"
         Tab(4).Control(4)=   "btntabcargocancel"
         Tab(4).Control(5)=   "btntabcargoalt"
         Tab(4).Control(6)=   "btntabcargosave"
         Tab(4).Control(7)=   "btntabcargodel"
         Tab(4).Control(8)=   "btntabcargoadd"
         Tab(4).Control(9)=   "cmbtabpesquisacargo"
         Tab(4).Control(10)=   "frmtabcargopesquisa"
         Tab(4).Control(11)=   "dtatabcargos"
         Tab(4).ControlCount=   12
         TabCaption(5)   =   "Níveis"
         TabPicture(5)   =   "Cadcli.frx":05BD
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txttabdescnivel"
         Tab(5).Control(1)=   "txttabcodnivel"
         Tab(5).Control(2)=   "btntabnivelcancel"
         Tab(5).Control(3)=   "btntabnivelalt"
         Tab(5).Control(4)=   "btntabnivelsave"
         Tab(5).Control(5)=   "btntabniveldel"
         Tab(5).Control(6)=   "btntabniveladd"
         Tab(5).Control(7)=   "cmbtabpesquisanivel"
         Tab(5).Control(8)=   "frmtabnivelpesquisa"
         Tab(5).Control(9)=   "dtatabnivel"
         Tab(5).Control(10)=   "lblcodnivel"
         Tab(5).Control(11)=   "lbldescnivel"
         Tab(5).ControlCount=   12
         Begin VB.TextBox txttabmailcli 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cli_mail"
            DataSource      =   "dtacliente"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   40
            TabIndex        =   99
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox txttabfonecli 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cli_fone"
            DataSource      =   "dtacliente"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   18
            TabIndex        =   98
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox txttabendcli 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cli_end"
            DataSource      =   "dtacliente"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   40
            TabIndex        =   97
            Top             =   1200
            Width           =   4575
         End
         Begin VB.TextBox txttabnomecli 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cli_nome"
            DataSource      =   "dtacliente"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   40
            TabIndex        =   96
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox txttabcodcli 
            DataField       =   "cli_cod"
            DataSource      =   "dtacliente"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   6
            TabIndex        =   95
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txttabcodequipe 
            DataField       =   "eq_mat"
            DataSource      =   "dtatabeq"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   6
            TabIndex        =   94
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txttabnomequipe 
            BackColor       =   &H00FFFFFF&
            DataField       =   "eq_nome"
            DataSource      =   "dtatabeq"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   40
            TabIndex        =   93
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox txttabendequipe 
            BackColor       =   &H00FFFFFF&
            DataField       =   "eq_end"
            DataSource      =   "dtatabeq"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   40
            TabIndex        =   92
            Top             =   1200
            Width           =   4575
         End
         Begin VB.TextBox txttabfonequipe 
            BackColor       =   &H00FFFFFF&
            DataField       =   "eq_fone"
            DataSource      =   "dtatabeq"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   18
            TabIndex        =   91
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox txttabmailequipe 
            BackColor       =   &H00FFFFFF&
            DataField       =   "eq_email"
            DataSource      =   "dtatabeq"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   40
            TabIndex        =   90
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox txttabcargoequipe 
            DataField       =   "eq_cargo"
            DataSource      =   "dtatabeq"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -71040
            MaxLength       =   2
            TabIndex        =   89
            Top             =   480
            Width           =   495
         End
         Begin VB.Data dtaproblemas 
            Connect         =   "Access"
            DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\hpdsk.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74760
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Tipo_problema"
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txttabsolcont 
            DataField       =   "sol_conteudo"
            DataSource      =   "dtasoluções"
            Enabled         =   0   'False
            Height          =   495
            Left            =   -70080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   88
            Top             =   1920
            Width           =   6135
         End
         Begin VB.Data dtasoluções 
            Connect         =   "Access"
            DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\hpdsk.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74760
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Solucao"
            Top             =   2160
            Width           =   2775
         End
         Begin VB.TextBox txttabsoldesc 
            DataField       =   "sol_desc"
            DataSource      =   "dtasoluções"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -70080
            TabIndex        =   87
            Top             =   1440
            Width           =   6135
         End
         Begin VB.Data dtatabeq 
            Connect         =   "Access"
            DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\hpdsk.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -67560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Equipe"
            Top             =   1860
            Width           =   3135
         End
         Begin VB.CommandButton btntabprobadd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   -70320
            TabIndex        =   86
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton btntabprobalt 
            Caption         =   "Editar"
            Height          =   375
            Left            =   -67680
            TabIndex        =   85
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton btntabprobdel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   -66360
            TabIndex        =   84
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton btntabsoladd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   -70320
            TabIndex        =   83
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton btntabsolsave 
            Caption         =   "Salvar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69000
            TabIndex        =   82
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton btntabsolalt 
            Caption         =   "Editar"
            Height          =   375
            Left            =   -67680
            TabIndex        =   81
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton btntabeqadd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   -74760
            TabIndex        =   80
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton btntabeqdel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   -70560
            TabIndex        =   79
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabeqsave 
            Caption         =   "Salvar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -73440
            TabIndex        =   78
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabcliadd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   -74760
            TabIndex        =   77
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton btntabclidel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   -70560
            TabIndex        =   76
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabclialt 
            Caption         =   "Editar"
            Height          =   375
            Left            =   -72000
            TabIndex        =   75
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Data dtacliente 
            Connect         =   "Access"
            DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\hpdsk.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -67560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Clientes"
            Top             =   1860
            Width           =   3135
         End
         Begin VB.CommandButton btntabclisave 
            Caption         =   "Salvar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -73440
            TabIndex        =   74
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Frame frmpesquisacli 
            Caption         =   "Pesquisa :"
            Height          =   855
            Left            =   -67560
            TabIndex        =   72
            Top             =   930
            Width           =   3135
            Begin VB.TextBox txttabclipesquisar 
               Height          =   285
               Left            =   120
               TabIndex        =   73
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.CommandButton btntabeqalt 
            Caption         =   "Editar"
            Height          =   375
            Left            =   -72000
            TabIndex        =   71
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Frame frmtabeqpesquisa 
            Caption         =   "Pesquisa :"
            Height          =   855
            Left            =   -67560
            TabIndex        =   69
            Top             =   930
            Width           =   3135
            Begin VB.TextBox txttabeqpesquisa 
               Height          =   285
               Left            =   120
               TabIndex        =   70
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.TextBox txttabnivcli 
            DataField       =   "cli_nivel"
            DataSource      =   "dtacliente"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -71040
            MaxLength       =   2
            TabIndex        =   68
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton btntabprobsave 
            Caption         =   "Salvar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69000
            TabIndex        =   67
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton btntabsoldel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   -66360
            TabIndex        =   66
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton btntabclicancel 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69120
            TabIndex        =   65
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabeqcancel 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69120
            TabIndex        =   64
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabprobcancel 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -65040
            TabIndex        =   63
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton btntabsolcancel 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -65040
            TabIndex        =   62
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txttabprobcod 
            DataField       =   "prob_cod"
            DataSource      =   "dtaproblemas"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            MaxLength       =   6
            TabIndex        =   61
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txttabprobdesc 
            DataField       =   "prob_tipo"
            DataSource      =   "dtaproblemas"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -70080
            MaxLength       =   20
            TabIndex        =   60
            Top             =   480
            Width           =   6135
         End
         Begin VB.ComboBox cmbtabpesquisacli 
            Height          =   315
            ItemData        =   "Cadcli.frx":05D9
            Left            =   -67560
            List            =   "Cadcli.frx":05EF
            TabIndex        =   59
            Top             =   600
            Width           =   3135
         End
         Begin VB.ComboBox cmbtabpesquisaeq 
            Height          =   315
            ItemData        =   "Cadcli.frx":0624
            Left            =   -67560
            List            =   "Cadcli.frx":063A
            TabIndex        =   58
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txttabsolcod 
            DataField       =   "sol_num"
            DataSource      =   "dtasoluções"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72240
            MaxLength       =   6
            TabIndex        =   57
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txttabsolprob 
            DataField       =   "sol_tipoprob"
            DataSource      =   "dtasoluções"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72840
            MaxLength       =   6
            TabIndex        =   56
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txttabdesccargo 
            BackColor       =   &H00FFFFFF&
            DataField       =   "cargo_des"
            DataSource      =   "dtatabcargos"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   30
            TabIndex        =   55
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox txttabcodcargo 
            DataField       =   "cargo_cod"
            DataSource      =   "dtatabcargos"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   6
            TabIndex        =   54
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton btntabcargocancel 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69120
            TabIndex        =   53
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabcargoalt 
            Caption         =   "Editar"
            Height          =   375
            Left            =   -72000
            TabIndex        =   52
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabcargosave 
            Caption         =   "Salvar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -73440
            TabIndex        =   51
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabcargodel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   -70560
            TabIndex        =   50
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabcargoadd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   -74760
            TabIndex        =   49
            Top             =   2400
            Width           =   1215
         End
         Begin VB.ComboBox cmbtabpesquisacargo 
            Height          =   315
            ItemData        =   "Cadcli.frx":066F
            Left            =   -67560
            List            =   "Cadcli.frx":0679
            TabIndex        =   48
            Top             =   600
            Width           =   3135
         End
         Begin VB.Frame frmtabcargopesquisa 
            Caption         =   "Pesquisa :"
            Height          =   855
            Left            =   -67560
            TabIndex        =   46
            Top             =   930
            Width           =   3135
            Begin VB.TextBox txttabcargopesquisa 
               Height          =   285
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.Data dtatabcargos 
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -67560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1860
            Width           =   3135
         End
         Begin VB.Data dtahistorico 
            Connect         =   "Access"
            DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\Hist.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Historico"
            Top             =   2520
            Width           =   2340
         End
         Begin VB.TextBox txthstchamacod 
            DataField       =   "hst_codchama"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   480
            MaxLength       =   6
            TabIndex        =   45
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthsttipoprob 
            DataField       =   "hst_tipoprob"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   44
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthstdataechama 
            DataField       =   "hst_dataechama"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   43
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthstdataschama 
            DataField       =   "hst_dataschama"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   42
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthstnomecli 
            DataField       =   "hst_nomecli"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   6480
            MaxLength       =   40
            TabIndex        =   41
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthstnomeeq 
            DataField       =   "hst_nomeeq"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   7680
            MaxLength       =   40
            TabIndex        =   40
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txthsthora 
            DataField       =   "hst_horachama"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   5280
            MaxLength       =   15
            TabIndex        =   39
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthstresp 
            DataField       =   "hst_nomeresp"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   8760
            MaxLength       =   40
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txthstsolsn 
            DataField       =   "hst_solsn"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   280
            Left            =   9840
            MaxLength       =   40
            TabIndex        =   37
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txthstn 
            DataField       =   "hst_historico"
            DataSource      =   "dtahistorico"
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   35
            Top             =   2520
            Width           =   855
         End
         Begin VB.Data dtatabnivel 
            Connect         =   "Access"
            DatabaseName    =   "D:\Arquivo\Faculdade\Programação de Computadores\3º ano\Programa\Cadastro\hpdsk.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -67560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Nivel"
            Top             =   1860
            Width           =   3135
         End
         Begin VB.Frame frmtabnivelpesquisa 
            Caption         =   "Pesquisa :"
            Height          =   855
            Left            =   -67560
            TabIndex        =   33
            Top             =   930
            Width           =   3135
            Begin VB.TextBox txttabpesquisanivel 
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.ComboBox cmbtabpesquisanivel 
            Height          =   315
            ItemData        =   "Cadcli.frx":0690
            Left            =   -67560
            List            =   "Cadcli.frx":069A
            TabIndex        =   32
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton btntabniveladd 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   -74760
            TabIndex        =   31
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton btntabniveldel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   -70560
            TabIndex        =   30
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabnivelsave 
            Caption         =   "Salvar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -73440
            TabIndex        =   29
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabnivelalt 
            Caption         =   "Editar"
            Height          =   375
            Left            =   -72000
            TabIndex        =   28
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton btntabnivelcancel 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69120
            TabIndex        =   27
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txttabcodnivel 
            DataField       =   "nivel_cod"
            DataSource      =   "dtatabnivel"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   6
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txttabdescnivel 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nivel_desc"
            DataSource      =   "dtatabnivel"
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73200
            MaxLength       =   30
            TabIndex        =   25
            Top             =   840
            Width           =   4575
         End
         Begin VB.ComboBox cmbhstfiltro 
            Height          =   315
            ItemData        =   "Cadcli.frx":06B1
            Left            =   7320
            List            =   "Cadcli.frx":06D3
            TabIndex        =   24
            Top             =   2520
            Width           =   3135
         End
         Begin VB.CommandButton btnhstfiltro 
            Height          =   375
            Left            =   10560
            Picture         =   "Cadcli.frx":073E
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Ordenar pelo campo selecionado"
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton btntabsolcopiar 
            Enabled         =   0   'False
            Height          =   375
            Left            =   -70920
            Picture         =   "Cadcli.frx":0B80
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Copiar Conteúdo para o Chamado"
            Top             =   2520
            Width           =   495
         End
         Begin MSDBGrid.DBGrid dbghst 
            Bindings        =   "Cadcli.frx":0FBE
            Height          =   1365
            Left            =   240
            OleObjectBlob   =   "Cadcli.frx":0FD9
            TabIndex        =   36
            Top             =   1080
            Width           =   10695
         End
         Begin VB.Label lbltabmailcli 
            Caption         =   "E-Mail :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   131
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lbltabfonecli 
            Caption         =   "Telefone :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   130
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lbltabendcli 
            Caption         =   "Endereço :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   129
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lbetabnomecli 
            Caption         =   "Nome :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   128
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lbltabcodcli 
            Caption         =   "Codigo do Cliente :"
            Height          =   255
            Left            =   -74880
            TabIndex        =   127
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblcodequipe 
            Caption         =   "Codigo do Atendente :"
            Height          =   255
            Left            =   -74880
            TabIndex        =   126
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblnomequipe 
            Caption         =   "Nome :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   125
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblendequipe 
            Caption         =   "Endereço :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   124
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblfoneequipe 
            Caption         =   "Telefone :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   123
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblmailquipe 
            Caption         =   "E-Mail :"
            Height          =   255
            Left            =   -74160
            TabIndex        =   122
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblcargoequipe 
            Caption         =   "Cargo :"
            Height          =   255
            Left            =   -72000
            TabIndex        =   121
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblcodproblema 
            Caption         =   "Tipos de Problema :"
            Height          =   255
            Left            =   -74760
            TabIndex        =   120
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Descrição :"
            Height          =   255
            Left            =   -71040
            TabIndex        =   119
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Soluções Conhecidas :"
            Height          =   255
            Left            =   -74760
            TabIndex        =   118
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Descrição :"
            Height          =   255
            Left            =   -71040
            TabIndex        =   117
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Conteúdo :"
            Height          =   255
            Left            =   -71040
            TabIndex        =   116
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lbltanivclii 
            Caption         =   "Nível :"
            Height          =   255
            Left            =   -72000
            TabIndex        =   115
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lbltabcodcargo 
            Caption         =   "Codigo do Cargo :"
            Height          =   255
            Left            =   -74880
            TabIndex        =   114
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lbltabdesccargo 
            Caption         =   "Descrição :"
            Height          =   255
            Left            =   -74400
            TabIndex        =   113
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblhstcodchama 
            Caption         =   "Chamado"
            Height          =   255
            Left            =   720
            TabIndex        =   112
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblhstdatae 
            Caption         =   "Abertura"
            Height          =   255
            Left            =   3120
            TabIndex        =   111
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblhstdatas 
            Caption         =   "Fechamento"
            Height          =   255
            Left            =   4200
            TabIndex        =   110
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblhstprob 
            Caption         =   "Problema"
            Height          =   255
            Left            =   1920
            TabIndex        =   109
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblhstnomecli 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   6840
            TabIndex        =   108
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblhsteq 
            Caption         =   "Atendente"
            Height          =   255
            Left            =   7800
            TabIndex        =   107
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblhsthora 
            Caption         =   "Período"
            Height          =   255
            Left            =   5520
            TabIndex        =   106
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblhstsolsn 
            Caption         =   "Solucionado"
            Height          =   255
            Left            =   9960
            TabIndex        =   105
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblresp 
            Caption         =   "Responsável"
            Height          =   255
            Left            =   8880
            TabIndex        =   104
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lbldescnivel 
            Caption         =   "Descrição :"
            Height          =   255
            Left            =   -74400
            TabIndex        =   103
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblcodnivel 
            Caption         =   "Codigo do Nível :"
            Height          =   255
            Left            =   -74880
            TabIndex        =   102
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lbltabnivcli 
            Height          =   255
            Left            =   -70440
            TabIndex        =   101
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lbltabcargoequipe 
            Height          =   255
            Left            =   -70440
            TabIndex        =   100
            Top             =   480
            Width           =   2775
         End
      End
      Begin VB.Label lblcodordem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   240
         TabIndex        =   149
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lbldataent 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data de Entrada :"
         Height          =   255
         Left            =   2520
         TabIndex        =   148
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbldatafin 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data de Finalização :"
         Height          =   255
         Left            =   2520
         TabIndex        =   147
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblcodcli 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo do Cliente :"
         Height          =   255
         Left            =   5640
         TabIndex        =   146
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblnomecli 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nome :"
         Height          =   255
         Left            =   5640
         TabIndex        =   145
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbltpchama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Problema :"
         Height          =   255
         Left            =   240
         TabIndex        =   144
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblsolchama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solução :"
         Height          =   255
         Left            =   240
         TabIndex        =   143
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lbldescpchamado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descrição do Chamado :"
         Height          =   255
         Left            =   240
         TabIndex        =   142
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblatpchama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Atendido Por:"
         Height          =   255
         Left            =   5640
         TabIndex        =   141
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblrpchama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Repassado Para:"
         Height          =   255
         Left            =   5640
         TabIndex        =   140
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblspchama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solucionado Por :"
         Height          =   255
         Left            =   5640
         TabIndex        =   139
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblhoraschama 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quantidade de Horas utilizadas: "
         Height          =   255
         Left            =   5640
         TabIndex        =   138
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label lblnivelchamacli 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8160
         TabIndex        =   137
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblchkadmin 
         Caption         =   "1"
         Height          =   255
         Left            =   10680
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblchkuser 
         Caption         =   "1"
         Height          =   255
         Left            =   10920
         TabIndex        =   135
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbluseraltcads 
         Caption         =   "1"
         Height          =   255
         Left            =   11400
         TabIndex        =   134
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbluserprobs 
         Caption         =   "1"
         Height          =   255
         Left            =   11640
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblusername 
         Caption         =   "nome"
         Height          =   255
         Left            =   10680
         TabIndex        =   132
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Menu menu_arq 
      Caption         =   "&Arquivo"
      HelpContextID   =   1
      Index           =   1
      Begin VB.Menu menu_chama_inc 
         Caption         =   "I&ncluir"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu menu_chama_alt 
         Caption         =   "&Editar"
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu menu_chama_del 
         Caption         =   "Exclu&ir"
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu menu_chama_save 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu menu_chama_cancel 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuseparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_logoff 
         Caption         =   "Efetuar Log off do usuário atual"
         Index           =   1
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuseparador2_1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_sair 
         Caption         =   "Sair"
         Index           =   6
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu menu_ferramentas 
      Caption         =   "&Ferramentas"
      Index           =   1
      Begin VB.Menu menu_toolbar 
         Caption         =   "&Barra de Ferramentas"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu menu_reindex 
         Caption         =   "&Reindexar Registros"
         Index           =   1
         Begin VB.Menu menu_reindex_chamado 
            Caption         =   "Chamados"
            Index           =   1
         End
         Begin VB.Menu menu_reindex_Cliente 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu menu_reindex_equipe 
            Caption         =   "Equipe"
            Index           =   1
         End
         Begin VB.Menu menu_reindex_Probsol 
            Caption         =   "Problemas e Soluções"
            Index           =   1
         End
         Begin VB.Menu menu_reindex_Cargo 
            Caption         =   "Cargos"
            Index           =   1
         End
         Begin VB.Menu menu_reindex_nivel 
            Caption         =   "Níveis"
            Index           =   1
         End
      End
      Begin VB.Menu mnu_horaatual 
         Caption         =   "&Hora Atual"
         Index           =   1
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu_pesquisa_chama 
         Caption         =   "&Pesquisa de Chamados"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuseparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_hidecadastros 
         Caption         =   "Esconder Cadastros"
         Index           =   1
         Begin VB.Menu mnu_escondercadastros 
            Caption         =   "Sim"
            Index           =   1
         End
         Begin VB.Menu mnu_mostrarcadastros 
            Caption         =   "Não"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mnuseparador4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_hstatualizasim 
         Caption         =   "&Atualizar histório automaticamente"
         Index           =   1
         Begin VB.Menu mnu_hsttimesim 
            Caption         =   "Sim"
            Index           =   1
         End
         Begin VB.Menu mnu_hsttimenao 
            Caption         =   "Não"
            Checked         =   -1  'True
            Index           =   1
         End
      End
   End
   Begin VB.Menu menu_admin 
      Caption         =   "A&dministrador"
      Index           =   1
      Begin VB.Menu menu_admin_users 
         Caption         =   "&Usuários e Senhas"
         Index           =   1
      End
      Begin VB.Menu menu_limparhst 
         Caption         =   "&Limpar Histórico"
         Index           =   1
      End
      Begin VB.Menu menu_backup 
         Caption         =   "Backup"
         Index           =   1
         Begin VB.Menu menu_grava_backup 
            Caption         =   "Gravar"
            Index           =   1
         End
         Begin VB.Menu menu_recupera_backup 
            Caption         =   "Recuperar"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnu_relat 
      Caption         =   "&Relatórios"
      Index           =   1
      Begin VB.Menu mnu__print_relcli 
         Caption         =   "Relatório de Clientes"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_print_releq 
         Caption         =   "Relatório de Equipes"
         Index           =   2
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu_print_relchama 
         Caption         =   "Relatório de Chamados"
         Index           =   3
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu_print_ordem 
         Caption         =   "Ordem de Serviço"
         Index           =   4
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnu_ajuda 
      Caption         =   "Aj&uda"
      Index           =   1
      Begin VB.Menu mnu_topicos 
         Caption         =   "Help Desk TOP - Tópicos de Ajuda"
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu_sobre 
         Caption         =   "Sobre o sistema - Help Desk TOP"
         Index           =   2
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim aux As Single, conttime As Single, contalt As String, timeligado As String
Dim hstnum As Single, chamanum As Single, clinum As Single, eqnum As Single, cargonum As Single, nivelnum As Single
Dim tipoentra As Single, tipoentracli As Single, tipoentraeq As Single, tipoentracargo As Single, tipoentranivel As Single
Dim tipoentraprob As Single, tipoentrasol As Single

Private Sub Form_Load()
    On Error GoTo trataerros
    timeligado = 2
    tmrcontrol.Interval = 10000
    dtacliente.DatabaseName = "hpdsk.mdb"
    dtacliente.RecordSource = "Clientes"
    dtacliente.Exclusive = False
    dtatabeq.DatabaseName = "hpdsk.mdb"
    dtatabeq.RecordSource = "Equipe"
    dtatabeq.Exclusive = False
    dtaproblemas.DatabaseName = "hpdsk.mdb"
    dtaproblemas.RecordSource = "Tipo_problema"
    dtaproblemas.Exclusive = False
    dtasoluções.DatabaseName = "hpdsk.mdb"
    dtasoluções.RecordSource = "Solucao"
    dtasoluções.Exclusive = False
    dtamainfile.DatabaseName = "hpdsk.mdb"
    dtamainfile.RecordSource = "Chamados"
    dtamainfile.Exclusive = False
    dtatabcargos.DatabaseName = "hpdsk.mdb"
    dtatabcargos.RecordSource = "Cargo"
    dtatabcargos.Exclusive = False
    dtatabnivel.DatabaseName = "hpdsk.mdb"
    dtatabnivel.RecordSource = "nivel"
    dtatabnivel.Exclusive = False
    dtahistorico.DatabaseName = "hist.mdb"
    dtahistorico.RecordSource = "Historico"
    dtahistorico.Exclusive = False
    cmbtabpesquisacli.ListIndex = 0
    cmbtabpesquisaeq.ListIndex = 0
    cmbtabpesquisacargo.ListIndex = 0
    cmbchamapesquisa.ListIndex = 0
    cmbtabpesquisanivel.ListIndex = 0
    cmbhstfiltro.ListIndex = 0
    ' -----------------------------------------
    dtasoluções.RecordSource = "select * from solucao where sol_tipoprob =" & Val(txttabprobcod.Text)
    dtasoluções.Refresh
    txttabsolprob.Text = Val(txttabprobcod.Text)
trataerros:
    Select Case Err.Number
        Case 91
    End Select
End Sub



'------------------------------------------------------------------------------------
'Comandos Referentes aos menus de controle
'------------------------------------------------------------------------------------
' referente ao menu arquivo
Private Sub menu_chama_inc_Click(Index As Integer)
    btnchamaadd_Click
End Sub
Private Sub menu_chama_alt_Click(Index As Integer)
    btnchamaalt_Click
End Sub
Private Sub menu_chama_del_Click(Index As Integer)
    btnchamadel_Click
End Sub
Private Sub menu_chama_save_Click(Index As Integer)
    btnchamasave_Click
End Sub
Private Sub menu_chama_cancel_Click(Index As Integer)
    btnchamacancel_Click
End Sub

Private Sub menu_grava_backup_Click(Index As Integer)
    frmBackup.Show
End Sub

' referente ao menu administrador
Private Sub menu_limparhst_Click(Index As Integer)
    On Error GoTo trataerros
    dtahistorico.Recordset.MoveFirst
    While dtahistorico.Recordset.EOF = False
        dtahistorico.Recordset.Delete
        dtahistorico.Recordset.MoveNext
    Wend
    dtahistorico.Recordset.MoveFirst
trataerros:
    Select Case Err.Number
    Case 3021
       MsgBox "O Histório está vazio", vbOKOnly, "Aviso de sistema"
    End Select
End Sub


Private Sub menu_recupera_backup_Click(Index As Integer)
    frmrestore.Show
End Sub

Private Sub mnu__print_relcli_Click(Index As Integer)
    rptclientes.Destination = 0
    rptclientes.RetrieveDataFiles
    rptclientes.Action = 1
End Sub

Private Sub mnu_print_releq_Click(Index As Integer)
    rpteq.Destination = 0
    rpteq.RetrieveDataFiles
    rpteq.Action = 1
End Sub

Private Sub mnu_print_relchama_Click(Index As Integer)
    rptchama.Destination = 0
    rptchama.RetrieveDataFiles
    rptchama.Action = 1
End Sub

Private Sub mnu_print_ordem_Click(Index As Integer)
    btnchamaprint_Click
End Sub

Private Sub menu_admin_users_Click(Index As Integer)
    frmbdusers.Show
End Sub

' menu ferramentas

Private Sub mnu_hsttimesim_Click(Index As Integer)
    timeligado = 1
    mnu_hsttimesim.Item(1).Checked = True
    mnu_hsttimenao.Item(1).Checked = False
End Sub
Private Sub mnu_hsttimenao_Click(Index As Integer)
    timeligado = 2
    mnu_hsttimenao.Item(1).Checked = True
    mnu_hsttimesim.Item(1).Checked = False
End Sub

Private Sub menu_reindex_chamado_Click(Index As Integer)
    If btnchamaadd.Enabled = True Then
        dtamainfile.RecordSource = "Select * from chamados Order by chama_num"
        dtamainfile.Refresh
    Else
        MsgBox "O arquivo de Chamados está aberto", vcokonly, "Não foi possível reindexar os registros"
    End If
End Sub

Private Sub menu_reindex_Cliente_Click(Index As Integer)
    If btntabcliadd.Enabled = True Then
        dtacliente.RecordSource = "Select * from clientes Order by cli_cod"
        dtacliente.Refresh
    Else
        MsgBox "O arquivo de Clientes está aberto", vcokonly, "Não foi possível reindexar os registros"
    End If
End Sub

Private Sub menu_reindex_equipe_Click(Index As Integer)
    If btntabeqadd.Enabled = True Then
        dtatabeq.RecordSource = "Select * from equipe Order by eq_mat"
        dtatabeq.Refresh
    Else
        MsgBox "O arquivo de Equipe está aberto", vcokonly, "Não foi possível reindexar os registros"
    End If
End Sub

Private Sub menu_reindex_Probsol_Click(Index As Integer)
    If btntabprobadd.Enabled = True Then
        dtaproblemas.RecordSource = "Select * from tipo_problema Order by prob_cod"
        dtaproblemas.Refresh
    Else
        MsgBox "O arquivo de Problemas está aberto", vcokonly, "Não foi possível reindexar os registros"
    End If
End Sub

Private Sub menu_reindex_Cargo_Click(Index As Integer)
    If btntabcargoadd.Enabled = True Then
        dtatabcargos.RecordSource = "Select * from cargo Order by cargo_cod"
        dtatabcargos.Refresh
    Else
        MsgBox "O arquivo de Cargos está aberto", vcokonly, "Não foi possível reindexar os registros"
    End If
End Sub

Private Sub menu_reindex_nivel_Click(Index As Integer)
    If btntabniveladd.Enabled = True Then
        dtatabnivel.RecordSource = "Select * from nivel Order by nivel_cod"
        dtatabnivel.Refresh
    Else
        MsgBox "O arquivo de Niveis está aberto", vcokonly, "Não foi possível reindexar os registros"
    End If
End Sub

Private Sub mnu_pesquisa_chama_Click(Index As Integer)
    If mnu_pesquisa_chama.Item(1).Checked = True Then
        frmpesquisachama.Enabled = False
        frmpesquisachama.Visible = False
        mnu_pesquisa_chama.Item(1).Checked = False
    Else
        frmpesquisachama.Enabled = True
        frmpesquisachama.Visible = True
        mnu_pesquisa_chama.Item(1).Checked = True
    End If
End Sub

Private Sub mnu_horaatual_Click(Index As Integer)
    MsgBox "A hora atual em seu sistema é " & Time, vbOKOnly, "Informação"
End Sub

Private Sub mnu_logoff_Click(Index As Integer)
    formentrasenha.Show
    Unload frmmain
End Sub

Private Sub menu_sair_Click(Index As Integer)
    Unload frmmain
End Sub

Private Sub menu_toolbar_Click(Index As Integer)
    frmtoolbar.Show
End Sub

Private Sub mnu_escondercadastros_Click(Index As Integer)
    tabmain.Visible = False
    mnu_escondercadastros.Item(1).Checked = True
    mnu_mostrarcadastros.Item(1).Checked = False
End Sub

Private Sub mnu_mostrarcadastros_Click(Index As Integer)
    tabmain.Visible = True
    mnu_mostrarcadastros.Item(1).Checked = True
    mnu_escondercadastros.Item(1).Checked = False
End Sub

Private Sub mnu_sobre_Click(Index As Integer)
    frmsobre.Show
End Sub

' Menu de ajuda

Private Sub mnu_topicos_Click(Index As Integer)
    cmndialog.HelpFile = "helpfile.hlp"
    cmndialog.HelpCommand = cdlHelpContents
    cmndialog.ShowHelp

End Sub

' Funções relativas apenas ao histórico

Private Sub tmrcontrol_Timer()
    On Error GoTo trataerros:
    If timeligado = 1 Then
        dtahistorico.Refresh
        dtahistorico.Recordset.MoveLast
    End If
trataerros:
    Select Case Err.Number
    End Select
End Sub

Private Sub dbghst_dblClick()
    If btnchamaadd.Enabled = True Then
        dtamainfile.Recordset.FindFirst "chama_num = " + txthstchamacod.Text
    End If
End Sub

Private Sub btnhstfiltro_Click()
    If cmbhstfiltro.ListIndex = 0 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_historico"
    End If
    If cmbhstfiltro.ListIndex = 1 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_codchama"
    End If
    If cmbhstfiltro.ListIndex = 2 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_tipoprob"
    End If
    If cmbhstfiltro.ListIndex = 3 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_dataechama"
    End If
    If cmbhstfiltro.ListIndex = 4 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_dataschama"
    End If
    If cmbhstfiltro.ListIndex = 5 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_horachama"
    End If
    If cmbhstfiltro.ListIndex = 6 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_nomecli"
    End If
    If cmbhstfiltro.ListIndex = 7 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_nomeeq"
    End If
    If cmbhstfiltro.ListIndex = 8 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_nomeresp"
    End If
    If cmbhstfiltro.ListIndex = 9 Then
        dtahistorico.RecordSource = "Select * from historico order by hst_solsn"
    End If
    dtahistorico.Refresh
End Sub

'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (CHAMADO)
'----------------------------------------------------------------
Private Sub btnchamaprint_Click()
    On Error GoTo trataerros
    rptchamado.Destination = 0
    rptchamado.RetrieveDataFiles
    rptchamado.ReplaceSelectionFormula "{chamados.chama_num} =" & txtcodchama
    rptchamado.Action = 1
trataerros:
    Select Case Err.Number
    Case 20515
        MsgBox "Selecione um chamado para imprimir", vbCritical, "Aviso do Sistema"
    End Select
End Sub

Private Sub txtcodchama_Change()
    On Error GoTo trataerros
    dtaproblemas.Recordset.FindFirst "prob_cod = " + dbctipoprob.BoundText
    dtasoluções.Refresh
trataerros:
    Select Case Err.Number
    End Select
End Sub

Private Sub btnchamaadd_Click()
    tipoentra = 1
    display_chama_controls
    dtamainfile.Refresh
    If dtamainfile.Recordset.EOF Then
        chamanum = 0
    Else
        dtamainfile.Recordset.MoveLast
        chamanum = Val(txtcodchama.Text) + 1
    End If
    dtamainfile.Recordset.AddNew
    btnchamasave.Enabled = True
    menu_chama_save.Item(1).Enabled = True
    btnchamacancel.Enabled = True
    menu_chama_cancel.Item(1).Enabled = True
    btnchamaadd.Enabled = False
    menu_chama_inc.Item(1).Enabled = False
    btnchamadel.Enabled = False
    menu_chama_del.Item(1).Enabled = False
    btnchamaalt.Enabled = False
    menu_chama_alt.Item(1).Enabled = False
    btntabsolcopiar.Enabled = True
    
    ' -----------------------------
    txtcodchama.Text = chamanum
    dbcequipeatp = lblusername.Caption
    txtdataechama.Text = Date
    txttabprobcod_change
    ' -----------------------------
    ' -----------------------------
    contalt = CStr(Hour(Time)) + ":" + CStr(Minute(Time))
    ' -----------------------------
    conttime = Time - TimeValue("00:00:00")
End Sub

Private Sub btnchamaalt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    tipoentra = 2
    display_chama_controls
    btnchamasave.Enabled = True
    menu_chama_save.Item(1).Enabled = True
    btnchamacancel.Enabled = True
    menu_chama_cancel.Item(1).Enabled = True
    btnchamaadd.Enabled = False
    menu_chama_inc.Item(1).Enabled = False
    btnchamadel.Enabled = False
    menu_chama_del.Item(1).Enabled = False
    btnchamaalt.Enabled = False
    menu_chama_alt.Item(1).Enabled = False
    btntabsolcopiar.Enabled = True
    ' -----------------------------
    contalt = CStr(Hour(Time)) + ":" + CStr(Minute(Time))
    ' -----------------------------
    conttime = Time - TimeValue(txthoraschama)
    ' -----------------------------
    txttabprobcod_change
    ' -----------------------------
    
trata_erros:
    Select Case Err.Number
        Case 13
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_chama_controls
                btnchamasave.Enabled = False
                menu_chama_save.Item(1).Enabled = False
                btnchamacancel.Enabled = False
                menu_chama_cancel.Item(1).Enabled = False
                btnchamaadd.Enabled = True
                menu_chama_inc.Item(1).Enabled = True
                btnchamadel.Enabled = True
                menu_chama_del.Item(1).Enabled = True
                btnchamaalt.Enabled = True
                menu_chama_alt.Item(1).Enabled = True
                btntabsolcopiar.Enabled = False
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_chama_controls
        End Select
End Sub

Private Sub btnchamacancel_Click()
    On Error GoTo trataerros
    If tipoentra = 1 Then
        dtamainfile.Recordset.CancelUpdate
    End If
    btnchamaadd.Enabled = True
    menu_chama_inc.Item(1).Enabled = True
    If lblchkuser = 0 Then
        btnchamadel.Enabled = True
        menu_chama_del.Item(1).Enabled = True
    End If
    btnchamaalt.Enabled = True
    menu_chama_alt.Item(1).Enabled = True
    btnchamasave.Enabled = False
    menu_chama_save.Item(1).Enabled = False
    btnchamacancel.Enabled = False
    menu_chama_cancel.Item(1).Enabled = False
    btntabsolcopiar.Enabled = False
    dtamainfile.Refresh
    hide_chama_controls
trataerros:
    Select Case Err.Number
    End Select
End Sub

Private Sub btnchamadel_Click() 'Botão - eliminar registro do chamado
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir este registro?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtamainfile.Recordset.Delete
        dtamainfile.Refresh
    End If
trata_erros:
    Select Case Err.Number
    Case 444
        dtamainfile.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub

Private Sub txtchamapesquisa_change()
    On Error GoTo trata_erros
    If cmbchamapesquisa.ListIndex = 0 Then
        dtamainfile.Recordset.FindFirst "chama_num = " + txtchamapesquisa.Text
    End If
    If cmbchamapesquisa.ListIndex = 1 Then
        dtamainfile.Recordset.FindFirst "chama_data_abre = '" + txtchamapesquisa.Text + "'"
    End If
    If cmbchamapesquisa.ListIndex = 2 Then
        dtamainfile.Recordset.FindFirst "chama_data_fecha = '" + txtchamapesquisa.Text + "'"
    End If
    If cmbchamapesquisa.ListIndex = 3 Then
        dtamainfile.Recordset.FindFirst "chama_qtd_horas = '" + txtchamapesquisa.Text + "'"
    End If
    If cmbchamapesquisa.ListIndex = 4 Then
        dtamainfile.Recordset.FindFirst "chama_atp = '" + txtchamapesquisa.Text + "'"
    End If
    If cmbchamapesquisa.ListIndex = 5 Then
        dtamainfile.Recordset.FindFirst "chama_repas = '" + txtchamapesquisa.Text + "'"
    End If
    If cmbchamapesquisa.ListIndex = 6 Then
        dtamainfile.Recordset.FindFirst "chama_sp = '" + txtchamapesquisa.Text + "'"
    End If
trata_erros:
    Select Case Err.Number
        Case 3077
        End Select
End Sub
Private Sub btnchamasave_Click() 'Botão - salvar registro ou atualizar modificações no registro de chamados
    On Error GoTo trata_erros
    If tipoentra = 2 Then
        dtamainfile.Recordset.Edit
    End If
    txthoraschama = conttime - Time
    If dbcequipesp.Text <> "" Then
       txtdataschama = Date
    End If
    '---Salva historico ----------------------------------
    dtahistorico.Refresh
    If dtahistorico.Recordset.EOF Then
            hstnum = 0
        Else
            dtahistorico.Recordset.MoveLast
            hstnum = Val(txthstn.Text) + 1
    End If
    dtahistorico.Recordset.AddNew
        txthstn = hstnum
        txthstchamacod.Text = txtcodchama.Text
        txthsttipoprob.Text = dbctipoprob.Text
        txthstdataechama.Text = txtdataechama.Text
        txthstdataschama.Text = txtdataschama.Text
        txthsthora = contalt + " - " + CStr(Hour(Time)) + ":" + CStr(Minute(Time))
        txthstnomecli.Text = txtnomecli.Text
        txthstnomeeq.Text = dbcequipeatp.Text
        txthstresp.Text = dbcequipesp.Text
        If txthstresp.Text = "" Then
            txthstsolsn.Text = "Não"
        Else
            txthstsolsn.Text = "Sim"
        End If
    
    dtahistorico.Recordset.Update
    contalt = ""
    dtahistorico.Recordset.MoveLast
    '----------------------------------------------------

    dtamainfile.Recordset.Update
    dtamainfile.Refresh
    btnchamaadd.Enabled = True
    menu_chama_inc.Item(1).Enabled = True
    If lblchkuser = 0 Then
        btnchamadel.Enabled = True
        menu_chama_del.Item(1).Enabled = False
    End If
    btnchamaalt.Enabled = True
    menu_chama_alt.Item(1).Enabled = True
    btnchamasave.Enabled = False
    menu_chama_save.Item(1).Enabled = False
    btnchamacancel.Enabled = False
    menu_chama_cancel.Item(1).Enabled = False
    btntabsolcopiar.Enabled = False
       
    hide_chama_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3426
            MsgBox "Um campo obrigatório não foi preenchido", vbOKOnly, "Não pode gravar o registro"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
        End Select
End Sub

Private Sub display_chama_controls() 'Função para habilitar as caixas de texto dos registros de chamados
    ' txtcodchama.Enabled = True
    txtdescchama.Enabled = True
    txtsolchama.Enabled = True
    txtcodcli.Enabled = True
    dbctipoprob.Enabled = True
    dbcequiperp.Enabled = True
    dbcequipesp.Enabled = True
    If tipoentra = 1 Then
        If lblchkuser = 0 Then
            dbcequipeatp.Enabled = True
        End If
    End If
    txtchamapesquisa.Enabled = False
    dtamainfile.Enabled = False
End Sub

Private Sub hide_chama_controls() 'Função para desabilitar as caixas de texto dos registros de chamados
    On Error GoTo trava:
'    txtcodchama.Enabled = False
    txtdescchama.Enabled = False
    txtsolchama.Enabled = False
    txtcodcli.Enabled = False
    dbctipoprob.Enabled = False
    dbcequipeatp.Enabled = False
    dbcequiperp.Enabled = False
    dbcequipesp.Enabled = False
    txtchamapesquisa.Enabled = True
    dtamainfile.Enabled = True
    dtamainfile.Recordset.MoveFirst
trava:
    Select Case Err.Number
    End Select
End Sub

Private Sub txtcodcli_Change()
    On Error GoTo trataerros
    If txtcodcli = "" Or txtcodcli = " " Then
        txtcodcli = ""
    End If
    dtacliente.Recordset.FindFirst "cli_cod = " + txtcodcli.Text
    If txtcodcli <> txttabcodcli And txtcodcli <> "" Then
        MsgBox "O codigo indicado para o cliente não está cadastrado", vbOKOnly, "aviso do sistema"
        txtcodcli = ""
        txtnomecli.Text = ""
        lblnivelchamacli = ""
    Else
        txtnomecli.Text = txttabnomecli.Text
        lblnivelchamacli = lbltabnivcli
    End If
trataerros:
    Select Case Err.Number
        Case 3077
            txtnomecli.Text = ""
            lblnivelchamacli = ""
    End Select
End Sub

Private Sub btntabsolcopiar_Click()
    txtsolchama = txttabsolcont
End Sub

Private Sub txtcodcli_LostFocus()
    On Error GoTo trataerros
    If txtcodcli.Text = "" Then
        dtacliente.Refresh
        dtacliente.Recordset.MoveFirst
        txtcodcli = txttabcodcli
    End If
trataerros:
    Select Case Err.Number
        Case 3021
            MsgBox "Não existe nenhum cliente cadastrado", vbCritical, "Não pode criar chamado"
            btnchamasave.Enabled = False
    End Select
End Sub




'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (clientes)
'----------------------------------------------------------------

Private Sub txttabnivcli_Change()
    On Error GoTo trataerros
    If txttabnivcli.Text = "" Then
        lbltabnivcli.Caption = ""
    Else
        dtatabnivel.Recordset.FindFirst "nivel_cod = " + txttabnivcli.Text
        lbltabnivcli.Caption = txttabdescnivel.Text
    End If
    If txttabnivcli <> txttabcodnivel And txttabnivcli <> "" Then
        MsgBox "O nível indicado para o cliente não está cadastrado", vbOKOnly, "aviso do sistema"
        txttabnivcli = ""
    End If
trataerros:
    Select Case Err.Number
        Case 3077
            txttabnivcli = ""
            dtatabnivel.Recordset.Edit
    End Select
End Sub

Private Sub btntabcliadd_Click() 'Botão - adicionar clientes
    display_cli_controls
    tipoentracli = 1
    dtacliente.Refresh
    If dtacliente.Recordset.EOF Then
        clinum = 0
    Else
        dtacliente.Recordset.MoveLast
        clinum = Val(txttabcodcli.Text) + 1
    End If
    dtacliente.Recordset.AddNew
    '-----------------------------------
    txttabcodcli.Text = clinum
    '-----------------------------------
    btntabclisave.Enabled = True
    btntabclicancel.Enabled = True
    btntabcliadd.Enabled = False
    btntabclidel.Enabled = False
    btntabclialt.Enabled = False
End Sub

Private Sub btntabclialt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    tipoentracli = 2
    display_cli_controls
    btntabclisave.Enabled = True
    btntabclicancel.Enabled = True
    btntabcliadd.Enabled = False
    btntabclidel.Enabled = False
    btntabclialt.Enabled = False
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_cli_controls
    End Select
End Sub

Private Sub btntabclicancel_Click()
    If tipoentracli = 1 Then
        dtacliente.Recordset.CancelUpdate
    End If
    btntabcliadd.Enabled = True
    btntabclidel.Enabled = True
    btntabclialt.Enabled = True
    btntabclisave.Enabled = False
    btntabclicancel.Enabled = False
    hide_cli_controls
    dtacliente.Refresh
End Sub

Private Sub btntabclidel_Click() 'Botão - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro do Cliente?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtacliente.Recordset.Delete
        dtacliente.Refresh
    Else

    End If
trata_erros:
    Select Case Err.Number
    Case 444
        dtacliente.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub



Private Sub txttabclipesquisar_change()
    On Error GoTo trata_erros
    If cmbtabpesquisacli.ListIndex = 0 Then
        dtacliente.Recordset.FindFirst "cli_cod = " + txttabclipesquisar.Text
    End If
    If cmbtabpesquisacli.ListIndex = 1 Then
        dtacliente.Recordset.FindFirst "cli_nome = '" + txttabclipesquisar.Text + "'"
    End If
    If cmbtabpesquisacli.ListIndex = 2 Then
        dtacliente.Recordset.FindFirst "cli_end = '" + txttabclipesquisar.Text + "'"
    End If
    If cmbtabpesquisacli.ListIndex = 3 Then
        dtacliente.Recordset.FindFirst "cli_fone = '" + txttabclipesquisar.Text + "'"
    End If
    If cmbtabpesquisacli.ListIndex = 4 Then
        dtacliente.Recordset.FindFirst "cli_mail = '" + txttabclipesquisar.Text + "'"
    End If
    If cmbtabpesquisacli.ListIndex = 5 Then
        dtacliente.Recordset.FindFirst "cli_nivel = " + txttabclipesquisar.Text
    End If
trata_erros:
    Select Case Err.Number
        Case 3077
        End Select
End Sub

Private Sub btntabclisave_Click() 'Botão - salvar registro ou atualizar modificações no registro de clientes
    On Error GoTo trata_erros
    If tipoentracli = 2 Then
        dtacliente.Recordset.Edit
    End If
    dtacliente.Recordset.Update
    dtacliente.Refresh
    btntabcliadd.Enabled = True
    btntabclidel.Enabled = True
    btntabclialt.Enabled = True
    btntabclisave.Enabled = False
    btntabclicancel.Enabled = False
    hide_cli_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub

Private Sub display_cli_controls() 'Função para habilitar as caixas de texto dos registros de clientes
'    txttabcodcli.Enabled = True
    txttabnomecli.Enabled = True
    txttabfonecli.Enabled = True
    txttabmailcli.Enabled = True
    txttabendcli.Enabled = True
    txttabnivcli.Enabled = True
    txttabclipesquisar.Enabled = False
    dtacliente.Enabled = False
End Sub
Private Sub hide_cli_controls() 'Função para desabilitar as caixas de texto dos registros de clientes
'    txttabcodcli.Enabled = False
    txttabnomecli.Enabled = False
    txttabfonecli.Enabled = False
    txttabmailcli.Enabled = False
    txttabendcli.Enabled = False
    txttabnivcli.Enabled = False
    txttabclipesquisar.Enabled = True
    dtacliente.Enabled = True
End Sub

Private Sub txttabnivcli_LostFocus()
    On Error GoTo trataerros
    If txttabnivcli.Text = "" Then
        dtatabnivel.Refresh
        dtatabnivel.Recordset.MoveFirst
        txttabnivcli = txttabcodnivel
    End If
trataerros:
    Select Case Err.Number
        Case 3021
            MsgBox "Não existe nenhum nível cadastrado", vbCritical, "Não pode cadastrar cliente"
            btntabclisave.Enabled = False
    End Select
End Sub

'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (Equipe)
'----------------------------------------------------------------
Private Sub txttabcargoequipe_Change()
    On Error GoTo trataerros
    If txttabcargoequipe = "" Or txttabcargoequipe = " " Then
        lbltabcargoequipe.Caption = ""
        txttabcargoequipe = ""
    Else
        dtatabcargos.Recordset.FindFirst "cargo_cod = " + txttabcargoequipe.Text
        lbltabcargoequipe.Caption = txttabdesccargo.Text
    End If
    If txttabcargoequipe <> txttabcodcargo And txttabcargoequipe <> "" Then
        MsgBox "O cargo indicado para o componente da equipe não está cadastrado", vbOKOnly, "aviso do sistema"
        txttabcargoequipe = ""
    End If
trataerros:
    Select Case Err.Number
    End Select
End Sub

Private Sub btntabeqadd_Click()
    display_equipe_controls
    tipoentraeq = 1
    dtatabeq.Refresh
    If dtatabeq.Recordset.EOF Then
        eqnum = 0
    Else
        dtatabeq.Recordset.MoveLast
        eqnum = Val(txttabcodequipe.Text) + 1
    End If
    dtatabeq.Recordset.AddNew
    '---------------------------------
    txttabcodequipe = eqnum
    '---------------------------------
    btntabeqsave.Enabled = True
    btntabeqcancel.Enabled = True
    btntabeqadd.Enabled = False
    btntabeqdel.Enabled = False
    btntabeqalt.Enabled = False
End Sub

Private Sub btntabeqalt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    display_equipe_controls
    tipoentraeq = 2
    btntabeqsave.Enabled = True
    btntabeqcancel.Enabled = True
    btntabeqadd.Enabled = False
    btntabeqdel.Enabled = False
    btntabeqalt.Enabled = False
    If txttabcodequipe = "" Then
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
        hide_equipe_controls
        btntabeqsave.Enabled = False
        btntabeqcancel.Enabled = False
        btntabeqadd.Enabled = True
        btntabeqdel.Enabled = True
        btntabeqalt.Enabled = True
    End If
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_equipe_controls
    End Select
End Sub

Private Sub btntabeqcancel_Click()
    On Error GoTo trataerros
    If tipoentraeq = 1 Then
        If btntabeqsave.Enabled = True Then
            dtatabeq.Recordset.CancelUpdate
        End If
    End If
    btntabeqadd.Enabled = True
    btntabeqdel.Enabled = True
    btntabeqalt.Enabled = True
    btntabeqsave.Enabled = False
    btntabeqcancel.Enabled = False
    hide_equipe_controls
    dtatabeq.Refresh
trataerros:
    Select Case Err.Number
    End Select
End Sub

Private Sub btntabeqdel_Click() 'Botão - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro do Cliente?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtatabeq.Recordset.Delete
        dtatabeq.Refresh
    Else

    End If
    dbcequipeatp.ReFill
    dbcequiperp.ReFill
    dbcequipesp.ReFill
trata_erros:
    Select Case Err.Number
    Case 444
        dtatabeq.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub

Private Sub txttabeqpesquisa_change()
    On Error GoTo trata_erros
    If cmbtabpesquisaeq.ListIndex = 0 Then
        dtatabeq.Recordset.FindFirst "eq_mat = " + txttabeqpesquisa.Text
    End If
    If cmbtabpesquisaeq.ListIndex = 1 Then
        dtatabeq.Recordset.FindFirst "eq_nome = '" + txttabeqpesquisa.Text + "'"
    End If
    If cmbtabpesquisaeq.ListIndex = 2 Then
        dtatabeq.Recordset.FindFirst "eq_end = '" + txttabeqpesquisa.Text + "'"
    End If
    If cmbtabpesquisaeq.ListIndex = 3 Then
        dtatabeq.Recordset.FindFirst "eq_fone = '" + txttabeqpesquisa.Text + "'"
    End If
    If cmbtabpesquisaeq.ListIndex = 4 Then
        dtatabeq.Recordset.FindFirst "eq_email = '" + txttabeqpesquisa.Text + "'"
    End If
    If cmbtabpesquisaeq.ListIndex = 5 Then
        dtatabeq.Recordset.FindFirst "eq_cargo = " + txttabeqpesquisa.Text
    End If
trata_erros:
    Select Case Err.Number
        Case 3077
        End Select
End Sub
Private Sub btntabeqsave_Click() 'Botão - salvar registro ou atualizar modificações no registro de clientes
    On Error GoTo trata_erros
    If tipoentraeq = 2 Then
        dtatabeq.Recordset.Edit
    End If
    dtatabeq.Recordset.Update
    dtatabeq.Refresh
    btntabeqadd.Enabled = True
    btntabeqdel.Enabled = True
    btntabeqalt.Enabled = True
    btntabeqsave.Enabled = False
    btntabeqcancel.Enabled = False
    hide_equipe_controls
    dbcequipeatp.ReFill
    dbcequiperp.ReFill
    dbcequipesp.ReFill
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub
Private Sub display_equipe_controls() 'Função para habilitar as caixas de texto dos registros de clientes
'    txttabcodequipe.Enabled = True
    txttabnomequipe.Enabled = True
    txttabfonequipe.Enabled = True
    txttabmailequipe.Enabled = True
    txttabendequipe.Enabled = True
    txttabcargoequipe.Enabled = True
    txttabeqpesquisa.Enabled = False
    dtatabeq.Enabled = False
End Sub

Private Sub hide_equipe_controls() 'Função para desabilitar as caixas de texto dos registros de clientes
'    txttabcodequipe.Enabled = False
    txttabnomequipe.Enabled = False
    txttabfonequipe.Enabled = False
    txttabmailequipe.Enabled = False
    txttabendequipe.Enabled = False
    txttabcargoequipe.Enabled = False
    txttabeqpesquisa.Enabled = True
    dtatabeq.Enabled = True
End Sub

Private Sub txttabcargoequipe_LostFocus()
    On Error GoTo trataerros
    If txttabcargoequipe.Text = "" Or txttabcargoequipe.Text = " " Then
        dtatabcargos.Refresh
        dtatabcargos.Recordset.MoveFirst
        txttabcargoequipe = txttabcodcargo
    End If
trataerros:
    Select Case Err.Number
        Case 3021
            MsgBox "Não existe nenhum cargo cadastrado", vbCritical, "Não pode cadastrar componente da equipe"
            btntabeqsave.Enabled = False
    End Select
End Sub

'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (Cargos)
'----------------------------------------------------------------

Private Sub btntabcargoadd_Click() 'Botão - adicionar clientes
    display_cargo_controls
    tipoentracargo = 1
    dtatabcargos.Refresh
    If dtatabcargos.Recordset.EOF Then
        cargonum = 0
    Else
        dtatabcargos.Recordset.MoveLast
        cargonum = Val(txttabcodcargo.Text) + 1
    End If
    dtatabcargos.Recordset.AddNew
    '---------------------------------
    txttabcodcargo = cargonum
    '---------------------------------
    btntabcargosave.Enabled = True
    btntabcargocancel.Enabled = True
    btntabcargoadd.Enabled = False
    btntabcargodel.Enabled = False
    btntabcargoalt.Enabled = False
End Sub

Private Sub btntabcargoalt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    display_cargo_controls
    tipoentracargo = 2
    btntabcargosave.Enabled = True
    btntabcargocancel.Enabled = True
    btntabcargoadd.Enabled = False
    btntabcargodel.Enabled = False
    btntabcargoalt.Enabled = False
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_cargo_controls
    End Select
End Sub

Private Sub btntabcargocancel_Click()
    If tipoentracargo = 1 Then
        dtatabcargos.Recordset.CancelUpdate
    End If
    btntabcargoadd.Enabled = True
    btntabcargodel.Enabled = True
    btntabcargoalt.Enabled = True
    btntabcargosave.Enabled = False
    btntabcargocancel.Enabled = False
    hide_cargo_controls
End Sub

Private Sub btntabcargodel_Click() 'Botão - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtatabcargos.Recordset.Delete
        dtatabcargos.Refresh
    End If
trata_erros:
    Select Case Err.Number
    Case 444
        dtatabcargos.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub


Private Sub txttabcargopesquisa_change()
    On Error GoTo trata_erros
    If cmbtabpesquisacargo.ListIndex = 0 Then
        dtatabcargos.Recordset.FindFirst "cargo_cod = " + txttabcargopesquisa.Text
    End If
    If cmbtabpesquisacargo.ListIndex = 1 Then
        dtatabcargos.Recordset.FindFirst "cargo_des = '" + txttabcargopesquisa.Text + "'"
    End If
trata_erros:
    Select Case Err.Number
        Case 3077
        End Select
End Sub

Private Sub btntabcargosave_Click() 'Botão - salvar registro ou atualizar modificações no registro de clientes
    On Error GoTo trata_erros
    If tipoentracargo = 2 Then
        dtatabcargos.Recordset.Edit
    End If
    dtatabcargos.Recordset.Update
    dtatabcargos.Refresh
    btntabcargoadd.Enabled = True
    btntabcargodel.Enabled = True
    btntabcargoalt.Enabled = True
    btntabcargosave.Enabled = False
    btntabcargocancel.Enabled = False
    hide_cargo_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub

Private Sub display_cargo_controls() 'Função para habilitar as caixas de texto dos registros de clientes
'    txttabcodcargo.Enabled = True
    txttabdesccargo.Enabled = True
    txttabcargopesquisa.Enabled = False
    dtatabcargos.Enabled = False
End Sub

Private Sub hide_cargo_controls() 'Função para desabilitar as caixas de texto dos registros de clientes
'    txttabcodcargo.Enabled = False
    txttabdesccargo.Enabled = False
    txttabcargopesquisa.Enabled = True
    dtatabcargos.Enabled = True
End Sub

'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (NÍVEIS)
'----------------------------------------------------------------

Private Sub btntabniveladd_Click() 'Botão - adicionar clientes
    display_nivel_controls
    tipoentranivel = 1
    dtatabnivel.Refresh
    If dtatabnivel.Recordset.EOF Then
        nivelnum = 0
    Else
        dtatabnivel.Recordset.MoveLast
        nivelnum = Val(txttabcodnivel.Text) + 1
    End If
    dtatabnivel.Recordset.AddNew
    '---------------------------------
    txttabcodnivel = nivelnum
    '---------------------------------
    btntabnivelsave.Enabled = True
    btntabnivelcancel.Enabled = True
    btntabniveladd.Enabled = False
    btntabniveldel.Enabled = False
    btntabnivelalt.Enabled = False
End Sub

Private Sub btntabnivelalt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    display_nivel_controls
    tipoentranivel = 2
    btntabnivelsave.Enabled = True
    btntabnivelcancel.Enabled = True
    btntabniveladd.Enabled = False
    btntabniveldel.Enabled = False
    btntabnivelalt.Enabled = False
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_nivel_controls
    End Select
End Sub

Private Sub btntabnivelcancel_Click()
    If tipoentranivel = 1 Then
        dtatabnivel.Recordset.CancelUpdate
    End If
    btntabniveladd.Enabled = True
    btntabniveldel.Enabled = True
    btntabnivelalt.Enabled = True
    btntabnivelsave.Enabled = False
    btntabnivelcancel.Enabled = False
    hide_nivel_controls
End Sub

Private Sub btntabniveldel_Click() 'Botão - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtatabnivel.Recordset.Delete
        dtatabnivel.Refresh
    End If
trata_erros:
    Select Case Err.Number
    Case 444
        dtatabnivel.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub



Private Sub txttabpesquisanivel_change()
    On Error GoTo trata_erros
    If cmbtabpesquisanivel.ListIndex = 0 Then
        dtatabnivel.Recordset.FindFirst "nivel_cod = " + txttabpesquisanivel.Text
    End If
    If cmbtabpesquisanivel.ListIndex = 1 Then
        dtatabnivel.Recordset.FindFirst "nivel_desc = '" + txttabpesquisanivel.Text + "'"
    End If
trata_erros:
    Select Case Err.Number
        Case 3077
        End Select
End Sub

Private Sub btntabnivelsave_Click() 'Botão - salvar registro ou atualizar modificações no registro de clientes
    On Error GoTo trata_erros
    If tipoentranivel = 2 Then
        dtatabnivel.Recordset.Edit
    End If
    dtatabnivel.Recordset.Update
    dtatabnivel.Refresh
    btntabniveladd.Enabled = True
    btntabniveldel.Enabled = True
    btntabnivelalt.Enabled = True
    btntabnivelsave.Enabled = False
    btntabnivelcancel.Enabled = False
    hide_nivel_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub

Private Sub display_nivel_controls() 'Função para habilitar as caixas de texto dos registros de clientes
'    txttabcodnivel.Enabled = True
    txttabdescnivel.Enabled = True
    txttabpesquisanivel.Enabled = False
    dtatabnivel.Enabled = False
End Sub

Private Sub hide_nivel_controls() 'Função para desabilitar as caixas de texto dos registros de clientes
'    txttabcodnivel.Enabled = False
    txttabdescnivel.Enabled = False
    txttabpesquisanivel.Enabled = True
    dtatabnivel.Enabled = True
End Sub


'----------------------------------------------------------------
'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (Tipo de Problema)
'----------------------------------------------------------------

Private Sub btntabprobadd_Click()
    tipoentraprob = 1
    If dtaproblemas.Recordset.EOF Then
        dtaproblemas.Recordset.AddNew
        txttabprobcod = "0"
        dtaproblemas.Recordset.Update
    End If
    dtaproblemas.Recordset.MoveLast
    aux = Val(txttabprobcod.Text) + 1
    dtaproblemas.Recordset.AddNew
    txttabprobcod.Text = aux
    display_problemas_controls
    btntabprobsave.Enabled = True
    btntabprobcancel.Enabled = True
    btntabprobadd.Enabled = False
    btntabprobdel.Enabled = False
    btntabprobalt.Enabled = False
End Sub

Private Sub btntabprobalt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    display_problemas_controls
    tipoentraprob = 2
    btntabprobsave.Enabled = True
    btntabprobcancel.Enabled = True
    btntabprobadd.Enabled = False
    btntabprobdel.Enabled = False
    btntabprobalt.Enabled = False
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_problemas_controls
    End Select
End Sub

Private Sub btntabprobcancel_Click()
    If tipoentraprob = 1 Then
        dtaproblemas.Recordset.CancelUpdate
    End If
    btntabprobadd.Enabled = True
    btntabprobdel.Enabled = True
    btntabprobalt.Enabled = True
    btntabprobsave.Enabled = False
    btntabprobcancel.Enabled = False
    hide_problemas_controls
End Sub

Private Sub btntabprobdel_Click() 'Botão - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro do Cliente?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtaproblemas.Recordset.Delete
        dtaproblemas.Refresh
    Else

    End If
    dbctipoprob.ReFill
trata_erros:
    Select Case Err.Number
    Case 444
        dtaproblemas.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub

Private Sub btntabprobsave_Click() 'Botão - salvar registro ou atualizar modificações no registro de clientes
    On Error GoTo trata_erros
    If tipoentraprob = 2 Then
        dtaproblemas.Recordset.Edit
    End If
    dtaproblemas.Recordset.Update
    dtaproblemas.Refresh
    btntabprobadd.Enabled = True
    btntabprobdel.Enabled = True
    btntabprobalt.Enabled = True
    btntabprobsave.Enabled = False
    btntabprobcancel.Enabled = False
    dtaproblemas.Recordset.MoveLast
    dbctipoprob.ReFill
    hide_problemas_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
            btntabprobadd.Enabled = True
            btntabprobdel.Enabled = True
            btntabprobalt.Enabled = True
            btntabprobsave.Enabled = False
            btntabprobcancel.Enabled = False
            hide_problemas_controls
            dtaproblemas.Recordset.MoveFirst
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub
Private Sub display_problemas_controls() 'Função para habilitar as caixas de texto dos registros de clientes
    txttabprobdesc.Enabled = True
    dtaproblemas.Enabled = False
End Sub

Private Sub hide_problemas_controls() 'Função para desabilitar as caixas de texto dos registros de clientes
    txttabprobdesc.Enabled = False
    dtaproblemas.Enabled = True
End Sub

'----------------------------------------------------------------
'Comandos Referentes ao cadastro, botões e caixas de controle do
'banco de dados (Solução)
'----------------------------------------------------------------

Private Sub txttabprobcod_change()
    dtasoluções.RecordSource = "select * from solucao where sol_tipoprob =" & Val(txttabprobcod.Text)
    dtasoluções.Refresh
    txttabsolprob.Text = Val(txttabprobcod.Text)
End Sub

    
Private Sub btntabsoladd_Click()
    tipoentrasol = 1
    dtasoluções.Refresh
    If dtasoluções.Recordset.EOF Then
       dtasoluções.Recordset.AddNew
       txttabsolcod = "0"
       txttabsolprob = Val(txttabprobcod.Text)
       dtasoluções.Recordset.Update
    End If
    dtasoluções.Recordset.MoveLast
    aux = Val(txttabsolcod.Text) + 1
    dtasoluções.Recordset.AddNew
    txttabsolprob.Text = Val(txttabprobcod.Text)
    txttabsolcod.Text = aux
    display_sol_controls
    btntabsolsave.Enabled = True
    btntabsolcancel.Enabled = True
    btntabsoladd.Enabled = False
    btntabsoldel.Enabled = False
    btntabsolalt.Enabled = False
End Sub

Private Sub btntabsolalt_Click() 'Botão - editar registro de clientes
    On Error GoTo trata_erros
    display_sol_controls
    tipoentrasol = 2
    btntabsolsave.Enabled = True
    btntabsolcancel.Enabled = True
    btntabsoladd.Enabled = False
    btntabsoldel.Enabled = False
    btntabsolalt.Enabled = False
trata_erros:
    Select Case Err.Number
        Case 3021
            MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
            hide_sol_controls
         Case 3260
            MsgBox "Este arquivo está sendo usado por outro usuário", vbOKOnly, "Problemas de rede"
            hide_sol_controls
    End Select
End Sub

Private Sub btntabsolcancel_Click()
    If tipoentrasol = 1 Then
        dtasoluções.Recordset.CancelUpdate
    End If
    btntabsoladd.Enabled = True
    btntabsoldel.Enabled = True
    btntabsolalt.Enabled = True
    btntabsolsave.Enabled = False
    btntabsolcancel.Enabled = False
    hide_sol_controls
End Sub

Private Sub btntabsoldel_Click() 'Botão - eliminar registro do cliente
    On Error GoTo trata_erros
    resp = MsgBox("Deseja excluir o registro do Cliente?", vbOKCancel, "Aviso de sistema")
    If resp = vbOK Then
        dtasoluções.Recordset.Delete
        dtasoluções.Refresh
    Else

    End If
trata_erros:
    Select Case Err.Number
    Case 444
        dtasoluções.Recordset.MoveFirst
    Case 3021
        MsgBox "Registro vazio ou inexistente", vbOKOnly, "Registro não encontrado"
    End Select
End Sub

Private Sub btntabsolsave_Click() 'Botão - salvar registro ou atualizar modificações no registro de clientes
    On Error GoTo trata_erros
    If tipoentrasol = 2 Then
        dtasoluções.Recordset.Edit
    End If
    dtasoluções.Recordset.Update
    dtasoluções.Refresh
    btntabsoladd.Enabled = True
    btntabsoldel.Enabled = True
    btntabsolalt.Enabled = True
    btntabsolsave.Enabled = False
    btntabsolcancel.Enabled = False
    hide_sol_controls
trata_erros:
    Select Case Err.Number
        Case 3058
            MsgBox "Não é possível gravar um registro sem um código", vbOKOnly, "Aviso de sistema"
        Case 3022
            MsgBox "Esse código já está cadastrado", vbOKOnly, "Não pode gravar o registro"
            btntabsoladd.Enabled = True
            btntabsoldel.Enabled = True
            btntabsolalt.Enabled = True
            btntabsolsave.Enabled = False
            btntabsolcancel.Enabled = False
            hide_sol_controls
            dtasoluções.Recordset.MoveFirst
        Case 3260
            MsgBox "Um outro usuário está gravando dados neste registro. Consulte o histórico ou aguarde e tente novamente... ", vbOKOnly, "Problemas de rede"
    End Select
End Sub
Private Sub display_sol_controls() 'Função para habilitar as caixas de texto dos registros de clientes
    txttabsoldesc.Enabled = True
    txttabsolcont.Enabled = True
    dtasoluções.Enabled = False
End Sub

Private Sub hide_sol_controls() 'Função para desabilitar as caixas de texto dos registros de clientes
    txttabsoldesc.Enabled = False
    txttabsolcont.Enabled = False
    dtasoluções.Enabled = True
End Sub

'------ FUNÇOES DE USUÁRIOS -------------------------------



