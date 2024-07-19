Attribute VB_Name = "ModPreencherWordESalvarPDF"
'Ativar em Referências o "Microsoft Word 16.0 Object Library"

Sub main()
    
    Const linhaTitulos              As Long = 9
    Const abaBase                   As String = "Individualizado"
    Const nomeArquivoWordModelo     As String = "MODELO MACRO.docx"
    Const celulaNomeLocal           As String = "B2"
    Const celulaEndereco            As String = "B3"
    
    Dim caminhoArquivoModelo        As String
    
    Dim aba                         As Worksheet
    Dim posColunaNome               As Long
    Dim posColunaRegistro           As Long
    Dim posColInicioApurado         As Long
    Dim posColFimApurado            As Long
    
    Dim nomeLocal                   As String
    Dim endereco                    As String
    Dim inicioApurado               As String
    Dim fimApurado                  As String
    
    Dim nome                        As String
    Dim registro                    As String
    
    Dim i                           As Long
    Dim ultimaLinha                 As Long
    
    ' Definir o caminho do arquivo modelo
    Let caminhoArquivoModelo = GetWorkbookPath(ThisWorkbook) & "\" & nomeArquivoWordModelo
    
    ' Verificar se o arquivo modelo existe
    If Dir(caminhoArquivoModelo) = "" Then
        MsgBox "Não localizado o arquivo " & caminhoArquivoModelo, vbCritical
        End
    End If

    ' Definir a aba
    Set aba = ThisWorkbook.Sheets(abaBase)
    
    ' Definir o nome do local
    Let nomeLocal = aba.Range(celulaNomeLocal)
    
    ' Definir o endereço
    Let endereco = aba.Range(celulaEndereco)
    
    ' Posição das colunas
    Let posColunaNome = posicaoColunaDoTitulo(aba, linhaTitulos, "Nome")
    Let posColunaRegistro = posicaoColunaDoTitulo(aba, linhaTitulos, "Registro")
    Let posColInicioApurado = posicaoColunaDoTitulo(aba, linhaTitulos, "Início Apurado")
    Let posColFimApurado = posicaoColunaDoTitulo(aba, linhaTitulos, "Fim Apurado")
    
    
    
    ' Definir a quantidade de linhas da tabela
    Let ultimaLinha = funcUltimaLinha(aba, posColunaNome)

    ' Criar objeto Word
    Set objWord = CreateObject("Word.Application")
    
    ' Deixar o word visivel
    objWord.Visible = True
        

    
    For i = linhaTitulos + 1 To ultimaLinha
    
        ' Definir o aruivo modelo
        Set arqModelo = objWord.Documents.Open(caminhoArquivoModelo)
    
        ' Definir o conteúdo do arquivo modelo
        Set conteudoDoc = arqModelo.Application.Selection
    
        ' Substituir dados do arquivo word modelo
        Let nome = aba.Cells(i, posColunaNome)
        Call substituirDadosDoWord(ByVal conteudoDoc, "[COLUNA A]", nome)
        
        Let registro = aba.Cells(i, posColunaRegistro)
        Call substituirDadosDoWord(ByVal conteudoDoc, "[COLUNA B]", registro)
        
        Call substituirDadosDoWord(ByVal conteudoDoc, "[B2]", nomeLocal)
        
        Call substituirDadosDoWord(ByVal conteudoDoc, "[B3]", endereco)
        
        Let inicioApurado = aba.Cells(i, posColInicioApurado)
        Call substituirDadosDoWord(ByVal conteudoDoc, "[COLUNA L]", inicioApurado)
        
        Let fimApurado = aba.Cells(i, posColFimApurado)
        Call substituirDadosDoWord(ByVal conteudoDoc, "[COLUNA M]", fimApurado)
        
        ' Salvar o novo arquivo preenchido
        arqModelo.SaveAs2 (GetWorkbookPath(ThisWorkbook) & "\" & i - linhaTitulos & " - " & nome & ".docx")
        
        ' Fechar o arquivo
        arqModelo.Close
        
        ' Criar um arquivo PDF utilizando o arquivo word salvo anteriormente
        Call SaveWordAsPDF((GetWorkbookPath(ThisWorkbook) & "\" & i - linhaTitulos & " - " & nome & ".docx"), _
                        (GetWorkbookPath(ThisWorkbook) & "\" & i - linhaTitulos & " - " & nome & ".pdf"))
            
        ' Excluir o arquivo word
        Call ExcluirArquivo(GetWorkbookPath(ThisWorkbook) & "\" & i - linhaTitulos & " - " & nome & ".docx")
        
    
    Next i
        

    objWord.Quit
    
    Set arqModelo = Nothing
    Set conteudoDoc = Nothing
    Set objWord = Nothing
    

    MsgBox "Documentos gerados!"

End Sub



Sub substituirDadosDoWord(ByVal conteudoDoc, ByVal textoDE As String, ByVal textoPARA)

        conteudoDoc.Find.Text = textoDE
        conteudoDoc.Find.Replacement.Text = textoPARA
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

End Sub




Sub SaveWordAsPDF(ByVal filePath As String, ByVal pdfPath As String)
    Dim wordApp     As Object
    Dim wordDoc     As Object
    
    
    ' Inicia o aplicativo Word
    Set wordApp = CreateObject("Word.Application")
    
    ' Abre o documento do Word
    Set wordDoc = wordApp.Documents.Open(filePath)
    
    ' Salva o documento como PDF
    wordDoc.SaveAs2 pdfPath, 17 ' 17 é o valor para wdFormatPDF
    
    ' Fecha o documento
    wordDoc.Close False
    ' Fecha o aplicativo Word
    wordApp.Quit
    
    ' Libera os objetos
    Set wordDoc = Nothing
    Set wordApp = Nothing

End Sub


Sub ExcluirArquivo(ByVal filePath As String)

    ' Verifica se o arquivo existe antes de tentar excluir
    If Dir(filePath) <> "" Then
        Kill filePath
    End If

End Sub






