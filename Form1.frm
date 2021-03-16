VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7320
      Top             =   6960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
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
Public isDBconected As Boolean



Private Sub bt_caixas_Click()
    
    mdl_connection.dbconect
    If isDBconected Then
        frm_caixas.Show
    End If
    

    

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

Private Sub Form_Load()
    isDBconected = False
End Sub

Private Sub frm_Menu_DblClick()

    MsgBox "Tem que abrir um caixa antes cara", vbCritical, "Aviso de Utilidade Pública"

End Sub

Public Function comandoSql(cmsSql)

End Function
