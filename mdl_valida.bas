Attribute VB_Name = "mdl_valida"
'modulo usado para anotar validações

Private Function validaData(data As String)
    On Error GoTo erroData
        nascimento = Format(data, "yyyy-mm-dd")
        validaData = True
        Exit Function
erroData:
validaData = False
Debug.Print "False"
End Function

