Attribute VB_Name = "mdl_conexao"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset

Function Connecta()

        On Error GoTo erroConexao
        
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Dim StringConexao As String
               
        StringConexao = "Driver={MySQL ODBC 3.51 Driver};Server= localhost;Port=3306;Database=" & frmConfigConexao.txtBanco.Text _
        & ";User=" & frmConfigConexao.txtUser.Text & ";Password=" & frmConfigConexao.txtSenha.Text _
        & ";Option=3;"
        
        cn.CursorLocation = adUseClient
        cn.ConnectionString = StringConexao
        cn.Open
        
        Exit Function
        
erroConexao:
        MsgBox "Ocorreu um erro na conexão!, tente novamente.", vbInformation, "Aviso"
        frmConfigConexao.Show vbModal
        
End Function


