VERSION 5.00
Begin VB.Form frm_Main 
   BackColor       =   &H80000016&
   Caption         =   "Lojinha - Versão VB6"
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
      Caption         =   "Sem conexão ao banco de dados - nenhuma ação será salva"
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
Public cmdSql As String
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public fld As ADODB.Field
Public SQL As String
Public libcon As ADODB.Connection
Public sqlstr As String


Private Sub bt_caixas_Click()
    frm_caixas.Show
    Module1.dbconect
    

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

Private Sub frm_Menu_DblClick()

    MsgBox "Tem que abrir um caixa antes cara", vbCritical, "Aviso de Utilidade Pública"

End Sub
Private Function conectadb2()
    MsgBox "teste"
End Function

Public Function conectaBD1()
    'TODO fazer a função para conectar no banco de dados, só copiei essa e não terminei de ver se ta certo'
 '   Set cnn = CreateObject("ADODB.Connection")

 '   cnn.Open "Driver={MariaDB ODBC 3.1.11 Driver};Server=localhost;UID=root;pwd=admin;database=lojinha; port=3306;option3"
'
 '   Set rs = CreateObject("ADODB.RecordSet")
'
 '   Set rs.ActiveConnection = cnn
  '  rs.Open "select * from Users"
'
 '   Ssql = "insert into Users (uname,upwd,uemail) values (Text1.text,Text3.text,Text2.text)"
  '  cnn.Execute Ssql
  On Error GoTo dbconnect_Exit

    
  Set libcon = New ADODB.Connection
  
  libcon.ConnectionString = "Driver={MariaDB ODBC 3.1.11 Driver};Server=localhost;UID=root;pwd=admin;db=lojinha; port=3306;option3"
  
  libcon.CursorLocation = adUseClient
  libcon.Open
  
dbconnect_Exit:
  MsgBox "Ocorreu um erro"
  
  
  

End Function

Public Function comandoSql(cmsSql)

End Function

Public Function conectaBD()
    Dim DBCon As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strName As String

    'Create a connection to the database
    Set DBCon = New ADODB.Connection
    DBCon.CursorLocation = adUseClient
    'This is a connectionstring to a local MySQL server
    DBCon.Open "server=localhost;pwd=admin;database=lojinha"

End Function

