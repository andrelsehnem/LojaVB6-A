VERSION 5.00
Begin VB.Form frm_novoCaixa 
   Caption         =   "Novo Caixa"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_valorCaixa 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "0"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txt_Descricao 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Novo caixa"
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton bt_cancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton bt_criar 
      Caption         =   "Criar"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lvl_valor 
      Caption         =   "Valor inicial em caixa"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lbl_descricao 
      Caption         =   "Descrição do novo caixa"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lbl_numCaixa 
      Caption         =   "Caixa número"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frm_novoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private desc As String
Private valor As Double


Private Sub bt_cancelar_Click()
    Unload Me

End Sub

Private Sub bt_criar_Click()
    Dim desc As String
    Dim valor As Double
    'TODO validar se os campos estão todos preenchidos, se não estiverem abrir um mensagem dizendo para preencher
    
    
    If txt_Descricao.Text = "" Then
        MsgBox "Preencha com um valor de descrição", vbCritical, "Atenção"
    Else
        desc = txt_Descricao.Text
         If txt_valorCaixa.Text = "" Then
            MsgBox "Preencha com um valor inicial para o caixa", vbCritical, "Atenção"
         Else
            valor = CDbl(txt_valorCaixa.Text)
            Unload Me
        End If
    End If
    
   
    
    'mdl_connection.ComandoSQL()
    
    'else
    'cria aviso que não ta certo

End Sub


Private Sub Form_Load()
    lbl_numCaixa.Caption = "Caixa número " & mdl_connection.SelectFrom("SELECT max(codigo) as codigo FROM caixas", "codigo") + 1
End Sub


