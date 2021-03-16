VERSION 5.00
Begin VB.Form frm_Main 
   BackColor       =   &H80000016&
   Caption         =   "Lojinha - Vers�o VB6"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19275
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
   ScaleHeight     =   7770
   ScaleWidth      =   19275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   15120
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton bt_CorFundo 
      Caption         =   "Alterar cor de Fundo"
      Height          =   615
      Left            =   16560
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton bt_fechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   18000
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
      Width           =   18975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Sem conex�o ao banco de dados - nenhuma a��o ser� salva"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7440
      Width           =   5295
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdSql As String


Private Sub bt_caixas_Click()
    frm_caixas.Show
    conectaBD

End Sub

Private Sub bt_CorFundo_Click()
    MsgBox "Esse bot�o vai alterar a cor de fundo do aplicativo", vbInformation, "Aviso de Utilidade P�blica"
End Sub

Private Sub bt_fechar_Click()
    Unload Me
End Sub

Private Sub bt_ValidCPF_Click()
    MsgBox "Esse bot�o vai validar um CPF", vbInformation, "Aviso de Utilidade P�blica"
End Sub

Private Sub frm_Menu_DblClick()

    MsgBox "Tem que abrir um caixa antes cara", vbCritical, "Aviso de Utilidade P�blica"

End Sub

Private Function conectaBD()
    'TODO fazer a fun��o para conectar no banco de dados, s� copiei essa e n�o terminei de ver se ta certo'
    Set cnn = CreateObject("ADODB.Connection")

    cnn.Open "driver={MySQL};server=localhost;pwd=admin;database=lojinha"

    Set Rs = CreateObject("ADODB.RecordSet")

    Set Rs.ActiveConnection = cnn
    Rs.Open "select * from Users"

    Ssql = "insert into Users (uname,upwd,uemail) values (Text1.text,Text3.text,Text2.text)"
    cnn.Execute Ssql

End Function

Private Function comandoSql(cmsSql)

End Function

