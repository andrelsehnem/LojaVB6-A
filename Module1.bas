Attribute VB_Name = "mdl_connection"
Public cmdSql As String
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public fld As ADODB.Field
Public SQL As String
Public libcon As ADODB.Connection
Public sqlstr As String


Sub dbconect()


        On Error GoTo erroConexao
        
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Dim StringConexao As String
               
        StringConexao = "Driver={MySQL ODBC 3.51 Driver};Server=" & frmConfigConexao.txtServidor.Text _
        & ";Port=" & frmConfigConexao.txtPorta.Text & ";Database=" & frmConfigConexao.txtBanco.Text _
        & ";User=" & frmConfigConexao.txtUser.Text & ";Password=" & frmConfigConexao.txtSenha.Text _
        & ";Option=3;"
        
        cn.CursorLocation = adUseClient
        cn.ConnectionString = StringConexao
        cn.Open
        
        Exit Function
dbconnect_Exit:
    MsgBox "Não Conectou"
  Exit Function
  
End Function


Sub dbComand(sqlstr)

    On Error GoTo dbcomandExit
    
    
    
    rs.Open sqlstr, libocon, adOpenKeyset, adLockOptimistic
    
    rs.Update
    rs.Close
    
    
    
    dbcomandExit
    MsgBox "Não executou o comando"
    Exit Sub
    
End Sub
