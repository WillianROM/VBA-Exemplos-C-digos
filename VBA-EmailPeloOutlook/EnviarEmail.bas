Attribute VB_Name = "EnviarEmail"
Option Explicit

Sub enviar_emaiil()
    Const Arquivo       As String = "C:\Users\Windows\Downloads\Projeto Outlook\RelatorioVendas.xlsx" 'Substitua pelo caminho de onde está RelatorioVendas.xlsx
    Dim objeto_outlook As Object
    Dim Email          As Object
    Dim texto1         As String
    Dim abaRelatorio    As Worksheet
    Dim NomeRelatorio   As String
    Dim WBk             As New Workbook
    Dim rng             As Range
    
    Set objeto_outlook = CreateObject("Outlook.Application")
    Set Email = objeto_outlook.CreateItem(0)
      
    Let NomeRelatorio = Right(Arquivo, (Len(Arquivo) - InStrRev(Arquivo, "\")))
    
    Set WBk = Workbooks(NomeRelatorio)
    
    Set abaRelatorio = WBk.Worksheets(1) 'Aba do relatório de onde será gerada uma imagem

    'Range na qual se gerará a imagem
    Set rng = abaRelatorio.Range("A1:E9")
    
    With Email
        .Display
        .To = "teste@testando.com.br"
        .CC = "teste1@testando.com.br"
        .BCC = "teste2@testando.com.br" 'copia oculta
        .Subject = "Aqui é o assunto do email"
    End With
    
    Let texto1 = "E aí <br><br>" & _
    "Olhe a imagem abaixo: <br><br>"

    Email.HTMLBody = texto1 & "<img src='C:\Users\Windows\Downloads\imgTeste.png'>" & _
    "<br><br>" & _
    RangetoHTML(rng) & _
    Email.HTMLBody 'Esse Email.HTMLBody é para colocar a assinatura que tinha antes de substituir o conteúdo
    
    
    Email.Attachments.Add (Arquivo)
    'Email.Send
    
End Sub
