Attribute VB_Name = "Mód_Executar_PFX"
Sub Executar_Arquivo_PFX()

    Dim objShell    As Object
    Dim caminho     As String
    Dim aba         As Worksheet
    
    
    Set aba = ThisWorkbook.Sheets("PRINCIPAL")

   
    Call listarArquivos(aba, "*.pfx")


    Set objShell = CreateObject("Shell.Application")
 
    caminho = Environ("userprofile") & "\Downloads\" & aba.Range("M1")

    objShell.Open (caminho)

    
    aba.Range("M1").Clear
    
End Sub


Function listarArquivos(ByVal aba As Worksheet, ByVal Extensao As String)
    Dim Arq


    'Variável armazena o local do arquivo
    Dim Pasta As String
    
    Let Pasta = Environ("userprofile") & "\Downloads\"
    
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
        aba.Range("M" & n) = Arq
        Arq = Dir
        n = n + 1
        
    Loop


End Function



