Attribute VB_Name = "Mod_00_Listar_Arquivos"
Function listarArquivos(Optional ByVal Extensao As String)
    Dim ABA     As Worksheet
    Dim Arq
    
    Set ABA = ThisWorkbook.Sheets("MACRO")

    'Variável armazena o local do arquivo
    Dim Pasta As String
    
    Let Pasta = ThisWorkbook.Path & "\"
    
    'Verifica se Existe a estensão do arquivo
    If Extensao = "" Then
        Extensao = "*"
    End If
    
    'Junta Pasta e Extensão
    Arq = Dir(Pasta & Extensao)
    
    'Informar o número da linha
    Dim n As Double
    Let n = 1
    
    'Verifica os arquivos até a variável Arq ficar vazia
    Do Until Arq = ""
        'Carrega o nome dos arquivos na célula
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
