
Option Explicit

Function listarArquivos(ByVal Aba As Worksheet, Optional ByVal extensao As String)
    Dim Pasta                       As String
    Dim QtdArquivosComAExtensao     As Long
    Dim n                           As Double
    Dim arrayNomesDosArquivos()
    Dim Arq

    Let QtdArquivosComAExtensao = CountFiles(extensao)

    ReDim arrayNomesDosArquivos(QtdArquivosComAExtensao - 1)

    'Variavel armazena o local do arquivo
    Let Pasta = ThisWorkbook.Path & "\"
    
    'Verifica se Existe a estensao do arquivo
    If extensao = "" Then extensao = "*"
    
    'Junta Pasta e Extens�o
    Let Arq = Dir(Pasta & extensao)
    
    'Informar o numero da linha
    
    Let n = 0
    
    'Verifica os arquivos até a variavel Arq ficar vazia
    Do Until Arq = ""
        'Carrega o nome dos arquivos na celula
        arrayNomesDosArquivos(n) = Arq
        
        Arq = Dir
        n = n + 1
        
    Loop

    listarArquivos = arrayNomesDosArquivos

End Function


Function CountFiles(ByVal extensao As String) As Long
    Dim xFolder         As String
    Dim xPath           As String
    Dim xCount          As Long
    Dim xFile           As String
    

    xFolder = ThisWorkbook.Path
    
    If xFolder = "" Then Exit Function
    
    xPath = xFolder & "\" & extensao
    xFile = Dir(xPath)
    
    Do While xFile <> ""
        xCount = xCount + 1
        xFile = Dir()
    Loop
    
    CountFiles = xCount
    
End Function
