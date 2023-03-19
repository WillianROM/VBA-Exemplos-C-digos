Attribute VB_Name = "Mod_LerInfoPDFs"
Option Explicit

'https://www.youtube.com/watch?v=7SUHaxOfYeQ&ab_channel=HashtagTreinamentos

Sub lerTextoPDF()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'Habilitar a referência Microsoft Word 16.0 Object Library
    Dim appWord             As New Word.Application
    Dim docWord             As Word.Document
    
    Dim nomeCompletoArq     As String
    Dim frasePDF            As Variant
    
    Dim arqPasta            As Variant
    Dim caminhoPasta        As String
    
    caminhoPasta = "C:\Users\Windows\Downloads\ResultadosDoExame\"
    arqPasta = Dir(caminhoPasta) 'Buscar arquivos dentro da pasta
    
    appWord.Visible = False
    
    Do While arqPasta <> ""
        nomeCompletoArq = caminhoPasta & arqPasta
    
        
        Set docWord = appWord.Documents.Open(nomeCompletoArq, False, True)
        
        'Debug.Print docWord.Sentences.Count

'        Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Value = arqPasta
'        Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = WorksheetFunction.Clean(WorksheetFunction.Trim(docWord.Sentences(2).Text))
        
        For Each frasePDF In docWord.Sentences
            Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Value = arqPasta
            Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = WorksheetFunction.Clean(frasePDF.Text)
        Next frasePDF
        
        
        arqPasta = Dir
    Loop
    
    docWord.Close False
    Set docWord = Nothing
    
    appWord.Quit
    Set appWord = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
