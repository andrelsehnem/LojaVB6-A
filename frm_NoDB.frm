VERSION 5.00
Begin VB.Form frm_NoDB 
   Caption         =   "Conexao Banco"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_fechar 
      Caption         =   "Fechar"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton bt_CriarBD 
      Caption         =   "Criar banco de dados"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton bt_ConectaBD 
      Caption         =   "Conectar ao Banco de Dados"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frm_NoDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isDBconected As Boolean
Public databaseName As String


Private Sub bt_ConectaBD_Click()
    
    Dbconect
    
    frm_Main.lbl_menu.Caption = "Conex?o estabelecida - Banco de dados " & databaseName
    frm_Main.Show
    Unload Me

End Sub

Private Sub Form_Load()
    databaseName = "lojinha"
     
    
End Sub
