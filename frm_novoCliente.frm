VERSION 5.00
Begin VB.Form frm_novoCliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6870
   ClientLeft      =   7980
   ClientTop       =   465
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_nascimento 
      Height          =   405
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txt_CPF 
      Height          =   375
      Left            =   240
      MaxLength       =   11
      TabIndex        =   3
      Text            =   " "
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txt_nomeCliente 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lbl_codigoCliente 
      Caption         =   "Código "
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lbl_nascimento 
      Caption         =   "Data de Nascimento*"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lbl_CPF 
      Caption         =   "CPF*"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lbl_nomeCliente 
      Caption         =   "Nome*"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frm_novoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As ADODB.Connection
Public rs As ADODB.Recordset


Private Sub Form_Load()
  
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    'cn.CursorLocation = adUseClient
    'cn.ConnectionString = mdl_connection.StringConexao
    'cn.CursorLocation = adUseClient
    
    
    Dim codigoCliente As Integer
    codigoCliente = 1
    
    Set rs = cn.Execute("SELECT max(codigo) as codigo FROM clientes")
    
    'SQL = "SELECT * FROM clientes ORDER BY codigo DESC LIMIT 1"
   ' rs.Open SQL, cn, adOpenStatic, adLockOptimistic
    'rs.AddNew
    
    If rs!codigo = 0 Then
        codigoCliente = 1
    Else
        codigoCliente = rs.Fields("codigo") + 1
    End If
    
    
    lbl_codigoCliente.Caption = rs!codigo
    
    
    
    lbl_codigoCliente.Caption = "Código " & codigoCliente
    
    Set rs = Nothing


End Sub

