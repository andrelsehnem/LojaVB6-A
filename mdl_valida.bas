Attribute VB_Name = "mdl_valida"
'modulo usado para anotar validações

Private Function validaData(data As String)
    On Error GoTo erroData
        nascimento = Format(data, "yyyy-mm-dd")
        validaData = True
        Exit Function
erroData:
MsgBox "Preencha a data corretamente"
validaData = False
End Function



'deixar essa função até não terminar a DLL

Public Function validaCPF(CPFtest As String)
    'SE RETORNAR TRUE É CPF VALIDO
    If CPFtest = "" Then
        validaCPF = False
    Dim cpf As String
    cpf = CPFtest
    Dim d1, d2 As Integer
    Dim soma As Integer
    soma = 0
    Dim digitado As String
    Dim calculado As String
    digitado = ""
    calculado = ""

    'Pesos para calcular o primeiro digito
    int[] peso1 = new int[] { 10, 9, 8, 7, 6, 5, 4, 3, 2 };
    'Pesos para calcular o segundo digito
    int[] peso2 = new int[] { 11, 10, 9, 8, 7, 6, 5, 4, 3, 2 };
    int[] n = new int[11];

    'Se o tamanho for < 11 entao retorna como inválido
    if (cpf.Length != 11)
        validaCPF = False

    try
        ' Quebra cada digito do CPF
        Dim contTrue As Integer
        contTrue = 0
        Dim nu As Integer
        For nu = 0 To 10
        
            n[nu] = Convert.ToInt32(cpf.Substring(nu, 1));
            if (nu > 0 && n[nu] == n[nu - 1])
             'validar se todos os numeors são iguais pt1
                contTrue++;
            
        
        if (contTrue >= 10)
         'validar se todos os numeros sao igais pt2
            validaCPF = False
        
    catch
        validaCPF = False
        
    'Calcula cada digito com seu respectivo peso
    Dim i As Integer
    For i = 0 To peso1.GetUpperBound(0)
        soma += (peso1[i] * Convert.ToInt32(n[i]));
    'Pega o resto da divisao
    int resto = soma % 11;

    if (resto == 1 || resto == 0)
        d1 = 0;
    Else
        d1 = 11 - resto;

    soma = 0;
    'Calcula cada digito com seu respectivo peso
    for (int i = 0; i <= peso2.GetUpperBound(0); i++)
        soma += (peso2[i] * Convert.ToInt32(n[i]));
    'Pega o resto da divisao
    resto = soma % 11;
    if (resto == 1 || resto == 0)
        d2 = 0;
    Else
        d2 = 11 - resto;
    calculado = d1.ToString() + d2.ToString();
    digitado = n[9].ToString() + n[10].ToString();
    'Se os ultimos dois digitos calculados bater com
    'os dois ultimos digitos do cpf entao é válido
    if (calculado == digitado)
        return (true);
    Else
        return (false);
