Attribute VB_Name = "Mód_CopiarDadosTXT"
Option Explicit


Sub copiarDadosDeUmArquivoTXTParaAPlanilha()
    Dim endereco As String
    Dim dadosTXT
    
    
    endereco = ThisWorkbook.Path & "\" & "teste.txt"
    
    dadosTXT = Importar_txt(endereco)
    
    Range(Cells(1, 1), Cells(1, 1).Offset(UBound(dadosTXT))).Value = dadosTXT

End Sub


Function Importar_txt(endereco As String)

    Dim arquivo
    
    Open endereco For Input As #1
        Importar_txt = Application.Transpose(Split(Input(LOF(1), 1), vbLf, -1))
    Close #1

End Function
