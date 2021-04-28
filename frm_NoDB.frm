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
    databaseName = "lojinha"
    Dbconect (databaseName)
        
    If isConected Then
    frm_Main.lbl_menu.Caption = "Conex?o estabelecida - Banco de dados " & databaseName
    frm_Main.Show
    
    Else
    MsgBox ("Conexao não estabelecida, verifique se já não esta conectado ou banco nao existe")
    End If
   
     
End Sub

Private Sub bt_CriarBD_Click()  'TODO verificar sempre se tem todas as tabelas aqui'
    databaseName = "lojinha"
    ServerConection
    On Error GoTo erroCriacaoBanco
    'Cria o banco'"
    ComandoSQL ("CREATE DATABASE " & frm_NoDB.databaseName)
    Dbconect (databaseName)
    'Tabela Caixa
    ComandoSQL ("CREATE TABLE caixas3 (codigo INT(4) NOT NULL AUTO_INCREMENT, descricao VARCHAR(30) NOT NULL COLLATE 'utf8_general_ci', valor DOUBLE NULL DEFAULT NULL, aberto TINYINT(1) NOT NULL DEFAULT '0',  ultimaAberto DATETIME NULL DEFAULT NULL,ultimoFechou DATETIME NULL DEFAULT NULL,appLanguageLog VARCHAR(10) NOT NULL DEFAULT 'ué' COLLATE 'utf8_general_ci',   PRIMARY KEY (codigo) USING BTREE)COLLATE='utf8_general_ci' ENGINE=InnoDB;")
    'Tabela Configurações
    ComandoSQL ("CREATE TABLE `configuracoes` (  `codigo` INT(4) NOT NULL AUTO_INCREMENT,    `descricao` VARCHAR(30) NOT NULL COLLATE 'utf8_general_ci', `config` VARCHAR(50) NOT NULL COLLATE 'utf8_general_ci',    `data` DATETIME NOT NULL,   `appLanguage` VARCHAR(10) NULL DEFAULT 'ué' COLLATE 'utf8_general_ci',  PRIMARY KEY (`codigo`) USING BTREE)COLLATE='utf8_general_ci'ENGINE=InnoDB AUTO_INCREMENT=4;")
    'Tabela log_login
    ComandoSQL ("CREATE TABLE `log_login` (  `codigo` INT(4) NOT NULL AUTO_INCREMENT,    `data` DATETIME NOT NULL,   `pc` VARCHAR(30) NULL DEFAULT NULL COLLATE 'utf8_general_ci',   `appLanguage` VARCHAR(10) NOT NULL DEFAULT 'ué' COLLATE 'utf8_general_ci',   PRIMARY KEY (`codigo`) USING BTREE)COLLATE='utf8_general_ci'ENGINE=InnoDB AUTO_INCREMENT=50;")
    'trigger configInsert
    ComandoSQL ("CREATE TRIGGER `configuracoesInsert` BEFORE INSERT ON `configuracoes` FOR EACH ROW SET NEW.data = NOW()")
    'trigger configUpdate
    ComandoSQL ("CREATE TRIGGER configUpdate BEFORE UPDATE ON configuracoes FOR EACH ROW SET new.data = NOW()")
    'trigger configDataHoje no login
    ComandoSQL ("CREATE TRIGGER `dataDeHoje` BEFORE INSERT ON `log_login` FOR EACH ROW SET NEW.data = NOW()")
    'Tabela clientes
    ComandoSQL ("CREATE TABLE `clientes` (   `codigo` INT(4) NOT NULL AUTO_INCREMENT,    `nome` VARCHAR(50) NOT NULL COLLATE 'utf8_general_ci', `cpf` VARCHAR(11) NOT NULL COLLATE 'utf8_general_ci',   `nascimento` DATE NOT NULL, `rua` VARCHAR(30) NULL DEFAULT NULL COLLATE 'utf8_general_ci',  `numeroRua` INT(5) NULL DEFAULT NULL,   `cidade` VARCHAR(20) NULL DEFAULT NULL COLLATE 'utf8_general_ci',   PRIMARY KEY (`codigo`) USING BTREE)COLLATE='utf8_general_ci' ENGINE=InnoDB;")
    
erroCriacaoBanco:
    MsgBox ("Ocorreu um erro, verifique se já não existe o banco e as tabelas")
    
    
End Sub

Private Sub bt_fechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub
