Attribute VB_Name = "M�d_Executar_PFX"
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


    'Vari�vel armazena o local do arquivo
    Dim Pasta As String
    
    Let Pasta = Environ("userprofile") & "\Downloads\"
    
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
        aba.Range("M" & n) = Arq
        Arq = Dir
        n = n + 1
        
    Loop


End Function



