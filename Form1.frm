VERSION 5.00
Begin VB.Form frm_Main 
   BackColor       =   &H80000016&
   Caption         =   "Lojinha - Versão VB6"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleMode       =   0  'User
   ScaleWidth      =   40193.05
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bt_caixas 
      Caption         =   "Caixa"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton bt_ValidCPF 
      Caption         =   "Validador de CPF"
      Height          =   615
      Left            =   11520
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton bt_CorFundo 
      Caption         =   "Alterar cor de Fundo"
      Height          =   615
      Left            =   12960
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton bt_fechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   14400
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame frm_Menu 
      BackColor       =   &H80000016&
      Caption         =   "Menu Inicial"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15375
      Begin VB.CommandButton bt_clientes 
         Caption         =   "Clientes"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame frm_operacional 
         BackColor       =   &H80000016&
         Caption         =   "Operacional"
         Height          =   6015
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label lbl_menu 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Sem conexão ao banco de dados - nenhuma ação será salva"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   7440
      Width           =   4350
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cmdSql As String
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public fld As ADODB.Field
Public SQL As String
Public libcon As ADODB.Connection
Public sqlstr As String



Private Sub bt_caixas_Click()
    frm_caixas.Show
    
End Sub

Private Sub bt_clientes_Click()
frm_clientes.Show
End Sub

Private Sub bt_CorFundo_Click()
    MsgBox "Esse botão vai alterar a cor de fundo do aplicativo", vbInformation, "Aviso de Utilidade Pública"
End Sub

Private Sub bt_fechar_Click()
    Unload Me
End Sub

Private Sub bt_ValidCPF_Click()
    MsgBox "Esse botão vai validar um CPF", vbInformation, "Aviso de Utilidade Pública"
End Sub


Public Sub Form_Load()
    
    ComandoSQL ("insert into log_login (pc,appLanguage) VALUES('pcAndre', 'VB6')")

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub frm_Menu_DblClick()

    MsgBox "Tem que abrir um caixa antes cara", vbCritical, "Aviso de Utilidade Pública"

End Sub



