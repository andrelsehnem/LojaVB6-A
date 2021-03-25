Attribute VB_Name = "mdl_connection"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Dim comandToDb As String
Public isConected As Boolean


Function Dbconect(nomeBanco As String)

        On Error GoTo erroConexao
        
        
        'nomeBanco = "lojinha"
       
        
        ServerConection
        comandToDb = "use " + nomeBanco + ";"
        
        ComandoSQL (comandToDb)
        
        
        frm_NoDB.isDBconected = True
        
        MsgBox "Conexão com o banco estabelecida", vbInformation, "Conectado"
        isConected = True
                
        Exit Function
        
erroConexao:
       MsgBox "Ocorreu um erro na conexão!, tente novamente.", vbInformation, "Aviso"
        
        
End Function


Function ComandoSQL(strcmd As String)
    
    rs.Open strcmd, cn, adOpenDynamic, adLockReadOnly
    
End Function

Function ServerConection()
    On Error GoTo erroConexao
        
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Dim StringConexao As String
               
        StringConexao = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;User=root;pwd=admin;port=3306;option3"
        
        cn.CursorLocation = adUseClient
        cn.ConnectionString = StringConexao
        cn.Open
        
        Exit Function
        
erroConexao:
       MsgBox "Ocorreu um erro na conexão!, tente novamente.", vbInformation, "Aviso"
        

End Function
