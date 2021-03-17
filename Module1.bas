Attribute VB_Name = "mdl_connection"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Dim comandToDb As String



Function Dbconect()

        On Error GoTo erroConexao
        
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Dim StringConexao As String
               
        StringConexao = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;User=root;pwd=admin;database=" & frm_NoDB.databaseName & "; port=3306;option3"
        
        cn.CursorLocation = adUseClient
        cn.ConnectionString = StringConexao
        cn.Open
        
        ComandoSQL ("insert into log_login (pc,appLanguage) VALUES('pcAndre', 'VB6')")
        frm_NoDB.isDBconected = True
        
        MsgBox "Conexão com o banco estabelecida", vbInformation, "Conectado"
 
        
        Exit Function
        
erroConexao:
       MsgBox "Ocorreu um erro na conexão!, tente novamente.", vbInformation, "Aviso"
        
        
End Function


Function ComandoSQL(strcmd As String)
    
    rs.Open strcmd, cn, adOpenDynamic, adLockReadOnly
    
End Function


Function CreateDB()
    
    
End Function
