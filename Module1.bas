Attribute VB_Name = "mdl_connection"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset


Function dbconect()



        On Error GoTo erroConexao
        
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Dim StringConexao As String
               
        StringConexao = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;User=root;pwd=admin;database=lojinha; port=3306;option3"
        
        cn.CursorLocation = adUseClient
        cn.ConnectionString = StringConexao
        cn.Open
        frm_Main.isDBconected = True
        
        Exit Function
        
erroConexao:
        MsgBox "Ocorreu um erro na conexão!, tente novamente.", vbInformation, "Aviso"
        
        
End Function


Sub dbComand(sqlstr)

  '  On Error GoTo dbcomandExit
    
    
    
 '   rs.Open sqlstr, libocon, adOpenKeyset, adLockOptimistic
    
 '   rs.Update
 '   rs.Close
    
    
    
 '   dbcomandExit
 '   MsgBox "Não executou o comando"
 '   Exit Sub
    
End Sub
