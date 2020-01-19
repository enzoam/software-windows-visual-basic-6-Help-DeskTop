VERSION 5.00
Begin VB.Form frmsobre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sobre o Sistema - Help Desk TOP"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnreg 
      Caption         =   "Registro"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton btnsobreok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Help Desk Solutions - 2000 - http://www.helpdesktop.cjb.net"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   6255
   End
   Begin VB.Label Label9 
      Caption         =   "LUIZ CRISTIANO CORDEIRO"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "LÚCIA MIEKO SUZUKI"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "JAMILE GEMA DE OLIVEIRA"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "ENZO AUGUSTO MARCHIORATO"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "APARECIDO RICARDO DE OLIVEIRA"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Analistas"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   600
      Picture         =   "sobre.frx":0000
      Top             =   240
      Width           =   5280
   End
   Begin VB.Label Label1 
      Caption         =   "Beta v 1.5  -  Novembro de 2000"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Projeto - Tecnologia em Processamento de Dados -  Programação de Computadores II "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   6135
   End
   Begin VB.Image Image2 
      Height          =   2130
      Left            =   1320
      Picture         =   "sobre.frx":36DA
      Top             =   1320
      Width           =   3990
   End
End
Attribute VB_Name = "frmsobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnreg_Click()
    MsgBox "Beta 1 - novembro de 2000 - v 1.5", btnokonly, "Registro"
End Sub

Private Sub btnsobreok_Click()
    Unload frmsobre
End Sub
