VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_novoCliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6045
   ClientLeft      =   7980
   ClientTop       =   465
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox msk_cpf 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      Format          =   "dd-mm-yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_nascimento 
      Bindings        =   "frm_novoCliente.frx":0000
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd-mm-yyyy"
      PromptChar      =   " "
   End
   Begin VB.CommandButton bt_cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton bt_gravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txt_estado 
      Height          =   375
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txt_cidade 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   7
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox txt_bairro 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   6
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txt_num 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txt_rua 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txt_nomeCliente 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lbl_estado 
      Caption         =   "Estado"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lbl_cidade 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lbl_bairro 
      Caption         =   "Bairro"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lbl_numero 
      Caption         =   "Numero"
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lbl_rua 
      Caption         =   "Rua"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lbl_codigoCliente 
      Caption         =   "Código "
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lbl_nascimento 
      Caption         =   "Data de Nascimento*"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lbl_CPF 
      Caption         =   "CPF*"
      Height          =   255
      Left            =   240
      TabIndex        =   11
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
Private nascimento As Date
Private cpf As String
Private validacao As Boolean




Private Sub bt_cancelar_Click()
    Unload Me
    
End Sub

Private Sub bt_gravar_Click()
    
    MsgBox validacao
    validacao = validaData
    'quando retorna true é que deu certo, então depois fazer as funções iguais da data pra os campos necessários e no final fazer um if para ver se validacao é TRUE ou FALS
    
    
    MsgBox validacao
    
    
End Sub

Private Sub Form_Load()
    validacao = False
    
    PegaCodCliente
    
End Sub


Function PegaCodCliente()
    
    Dim codigoCliente As Integer
    
    Dim tempCod As String
    
    tempCod = mdl_connection.SelectFrom("SELECT max(codigo) as codigo FROM clientes", "codigo")
    
    codigoCliente = CInt(tempCod)
    
    If codigoCliente = 0 Then
        codigoCliente = 1
    Else
        codigoCliente = codigoCliente + 1
    End If
        
    lbl_codigoCliente.Caption = "Codigo: " & codigoCliente


End Function

Public Function insereCliente()
    
    MsgBox ("INSERT INTO clientes (nome, cpf, nascimento, rua, numeroRua, bairro, cidade, estado) VALUES ('" & txt_nomeCliente.Text & "'," & cpf & ",'" & _
    nacimento & "','" & txt_rua.Text & "'," & txt_num.Text & ",'" & txt_bairro.Text & "','" & txt_cidade.Text & "','" & txt_estado.Text & "')")
    
    'ComandoSQL ("INSERT INTO clientes (nome, cpf, nascimento, rua, numeroRua, bairro, cidade, estado) VALUES ('" & txt_nomeCliente.Text & "'," & txt_CPF.Text & ",'" & _
    CDate(txt_nascimento.Text) & "','" & txt_rua.Text & "'," & CInt(txt_num.Text) & ",'" & txt_bairro.Text & "','" & txt_cidade.Text & "','" & txt_estado.Text & "')")


End Function

Private Function validaData()
    On Error GoTo erroData
        nascimento = Format(msk_nascimento.Text, "yyyy-mm-dd")
        validaData = True
        Exit Function
erroData:
MsgBox "Preencha a data de nascimento corretamente"
validaData = False
End Function
