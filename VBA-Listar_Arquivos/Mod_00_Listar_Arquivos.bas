Attribute VB_Name = "Mod_00_Listar_Arquivos"
Function listarArquivos(Optional ByVal Extensao As String)
    Dim ABA     As Worksheet
    Dim Arq
    
    Set ABA = ThisWorkbook.Sheets("MACRO")

    'Vari�vel armazena o local do arquivo
    Dim Pasta As String
    
    Let Pasta = ThisWorkbook.Path & "\"
    
    'Verifica se Existe a estens�o do arquivo
    If Extensao = "" Then
        Extensao = "*"
    End If
    
    'Junta Pasta e Extens�o
    Arq = Dir(Pasta & Extensao)
    
    'Informar o n�mero da linha
    Dim n As Double
    Let n = 1
    
    'Verifica os arquivos at� a vari�vel Arq ficar vazia
    Do Until Arq = ""
        'Carrega o nome dos arquivos na c�lula
        ABA.Range("M" & n) = Arq
        Arq = Dir
        n = n + 1
        
    Loop


End Function


Sub PesquisaArquivo()
    Dim ABA     As Worksheet
    
    Set ABA = ThisWorkbook.Sheets("MACRO")

    listarArquivos ("*.xlsx")

End Sub
