VERSION 5.00
Begin VB.Form frm_caixas 
   Caption         =   "Caixas Disponiveis"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8805
   LinkTopic       =   "frm_caixas"
   ScaleHeight     =   4890
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton bt_newCaixa 
      Caption         =   "Novo Caixa"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frm_caixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
