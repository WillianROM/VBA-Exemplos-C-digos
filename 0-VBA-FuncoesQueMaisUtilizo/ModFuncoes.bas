Attribute VB_Name = "ModFuncoes"
Option Explicit

Function VerificarSeArquivoExiste(ByVal caminhoDaPasta As String, ByVal nomeDoArquivo As String)
    
    ' Use a função Dir para verificar se o arquivo existe
    VerificarSeArquivoExiste = Dir(caminhoDaPasta & nomeDoArquivo) <> ""
    
End Function


Function posicaoLinhaDoTitulo(ByVal aba As Worksheet, linhaDoTitulo As Long, tituloALocalizar As String)
	Dim i           As Long
    	Dim qtdLinhas   As Long
    
    	Let qtdLinhas = WorksheetFunction.CountA(aba.Columns(linhaDoTitulo))

    	For i = 1 To qtdLinhas
        		If aba.Cells(i, linhaDoTitulo) = tituloALocalizar Then
            			posicaoLinhaDoTitulo = i
            			Exit For
        		End If
    	Next i

End Function


Function posicaoColunaDoTitulo(ByVal aba As Worksheet, linhaDoTitulo As Long, tituloALocalizar As String)
    	Dim i           As Long
    	Dim qtdColunas  As Long
    
    	Let qtdColunas = funcUltimaColuna(aba, linhaDoTitulo)

    	For i = 1 To qtdColunas
        		If aba.Cells(linhaDoTitulo, i) = tituloALocalizar Then
            			posicaoColunaDoTitulo = i
            			Exit For
        		End If
    	Next i

End Function


Function funcUltimaLinha(ByVal aba As Worksheet, ByVal posColuna As Long)

    	aba.Activate
    
    	' Definir a posição da última linha
    	aba.Cells(Rows.Count, posColuna).End(xlUp).Select
    	funcUltimaLinha = ActiveCell.Row
    
End Function


Function funcUltimaColuna(ByVal aba As Worksheet, ByVal linha As Long)

    	aba.Activate
    
    	' Definir a posição da última linha
    	aba.Cells(linha, Columns.Count).End(xlToLeft).Select
    	funcUltimaColuna = ActiveCell.Column
    
End Function


Function formatarCpfCnpj(ByVal txt_cpf_cnpj As String)

    	If Len(txt_cpf_cnpj) <= 11 Then
        f		ormatarCpfCnpj = Format(txt_cpf_cnpj, "000"".""000"".""000""-""00")
    	Else
        		formatarCpfCnpj = Format(txt_cpf_cnpj, "00"".""000"".""000""/""0000""-""00")
    	End If

End Function


Function listarArquivos(Optional ByVal extensao As String)
    	Dim Pasta                       As String
    	Dim QtdArquivosComAExtensao     As Long
    	Dim n                           As Double
    	Dim arrayNomesDosArquivos()
    	Dim Arq

	Let QtdArquivosComAExtensao = CountFiles(extensao)

    	ReDim arrayNomesDosArquivos(QtdArquivosComAExtensao - 1)

    	'Variável armazena o local do arquivo
    
    	Let Pasta = ThisWorkbook.Path & "\"
    
    	'Verifica se Existe a estensão do arquivo
    	If extensao = "" Then extensao = "*"
    
    	'Junta Pasta e Extensão
    	Let Arq = Dir(Pasta & extensao)
    
    	'Informar o número da linha
    
    	Let n = 0
    
    	'Verifica os arquivos até a variável Arq ficar vazia
    	Do Until Arq = ""
        		'Carrega o nome dos arquivos na célula
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


Public Function IsOutlookRunning() As Boolean
    	Dim olApp As Object
    
    	On Error Resume Next
    	Set olApp = GetObject(, "Outlook.Application")
    	On Error GoTo 0
    
    	If Not olApp Is Nothing Then
        		IsOutlookRunning = True
    	Else
        		IsOutlookRunning = False
    	End If
    
    	Set olApp = Nothing
End Function


Function CriarMatrizRemovendoRepeticoes(ByVal ws As Worksheet, ByVal posColuna As Long, ByVal linhaInicial As Long, ByVal linhaFinal As Long) As Variant

    	Dim matriz()        	As Variant
    	Dim coluna          	As Range
    	Dim i               	As Long
    	Dim j               	As Long
    	Dim k               	As Long
    	Dim valorAtual      	As Variant
    	Dim linhasMatriz    	As Long
   
   	 ' Selecionar a coluna com os dados
    	Set coluna = ws.Range(Cells(linhaInicial, posColuna).Address, Cells(linhaFinal, posColuna).Address)
   
    	' Inicializar a matriz com todos os valores da coluna
    	ReDim matriz(0 To coluna.Rows.Count - 1)
    	For i = 1 To coluna.Rows.Count
        		matriz(i - 1) = coluna.Cells(i, 1).Value
    	Next i
   
    	' Remover os valores repetidos da matriz

    '==================================================
    	Dim nPrimeira       	As Long
    	Dim nUltima         	As Long
    	Dim item            	As String
   
    	Dim arrTemp()	As String
    	Dim Coll            	As New Collection
   
     	'Obter a primeira e a última posição da matriz
     	nPrimeira = LBound(matriz)
     	nUltima = UBound(matriz)
     	ReDim arrTemp(nPrimeira To nUltima)
   
     	'Converter matriz em string
     	For i = nPrimeira To nUltima
        		arrTemp(i) = CStr(matriz(i))
     	Next i
       
     	'Preencher a coleção temporária
     	On Error Resume Next
     	For i = nPrimeira To nUltima
        		Coll.Add arrTemp(i), arrTemp(i)
     	Next i
     	Err.Clear
     	On Error GoTo 0
   
     	'Redimensionar a matriz
     	nUltima = Coll.Count + nPrimeira - 1
     	ReDim arrTemp(nPrimeira To nUltima)
       
     	'Preencher a matriz
     	For i = nPrimeira To nUltima
        		arrTemp(i) = Coll(i - nPrimeira + 1)
     	Next i
       
   
   
    '================================================================
   
     	' Redimensionar a matriz para remover as linhas vazias
    	ReDim Preserve arrTemp(0 To nUltima)
   
    	' Retornar a matriz
    	CriarMatrizRemovendoRepeticoes = arrTemp()
   
End Function


'=======================================================================================================================
' Funções criadas por outras pessoas
'=======================================================================================================================

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
   	'função para identificar se algum item corresponde ao conjunto de array
   	IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


Function emojiToChrW(emoji As String) As String

	Dim output 	As String
	Dim i		As Long

	For i = 1 To Len(emoji)

		output = output & "ChrW(" & AscW(Mid(emoji, i, 1)) & ") & "

	Next i

	output = Left(output, Len(output) - 3)

	emojiToChrW = output

End Function


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
   
End Function


Function GetWorkbookPath(Optional wb As Workbook)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Purpose:  Returns a workbook's physical path, even when they are saved in
    '           synced OneDrive Personal, OneDrive Business or Microsoft Teams folders.
    '           If no value is provided for wb, it's set to ThisWorkbook object instead.
    ' Author:   Ricardo Gerbaudo
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    GetWorkbookPath = wb.Path
    
    If InStr(1, wb.Path, "https://") <> 0 Then
        
        Const HKEY_CURRENT_USER = &H80000001
        Dim objRegistryProvider As Object
        Dim strRegistryPath As String
        Dim arrSubKeys()
        Dim strSubKey As Variant
        Dim strUrlNamespace As String
        Dim strMountPoint As String
        Dim strLocalPath As String
        Dim strRemainderPath As String
        Dim strLibraryType As String
    
        Set objRegistryProvider = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
        strRegistryPath = "SOFTWARE\SyncEngines\Providers\OneDrive"
        objRegistryProvider.EnumKey HKEY_CURRENT_USER, strRegistryPath, arrSubKeys
        
        For Each strSubKey In arrSubKeys
            objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "UrlNamespace", strUrlNamespace
            If InStr(1, wb.Path, strUrlNamespace) <> 0 Or InStr(1, strUrlNamespace, wb.Path) <> 0 Then
                objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "MountPoint", strMountPoint
                objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "LibraryType", strLibraryType
                
                If InStr(1, wb.Path, strUrlNamespace) <> 0 Then
                    strRemainderPath = Replace(wb.Path, strUrlNamespace, vbNullString)
                Else
                    GetWorkbookPath = strMountPoint
                    Exit Function
                End If
                
                'If OneDrive Personal, skips the GUID part of the URL to match with physical path
                If InStr(1, strUrlNamespace, "https://d.docs.live.net") <> 0 Then
                    If InStr(2, strRemainderPath, "/") = 0 Then
                        strRemainderPath = vbNullString
                    Else
                        strRemainderPath = Mid(strRemainderPath, InStr(2, strRemainderPath, "/"))
                    End If
                End If
                
                'If OneDrive Business, adds extra slash at the start of string to match the pattern
                strRemainderPath = IIf(InStr(1, strUrlNamespace, "my.sharepoint.com") <> 0, "/", vbNullString) & strRemainderPath
                
                strLocalPath = ""
                
                If (InStr(1, strRemainderPath, "/")) <> 0 Then
                    strLocalPath = Mid(strRemainderPath, InStr(1, strRemainderPath, "/"))
                    strLocalPath = Replace(strLocalPath, "/", "\")
                End If
                
                strLocalPath = strMountPoint & strLocalPath
                GetWorkbookPath = strLocalPath
                If Dir(GetWorkbookPath & "\" & wb.Name) <> "" Then Exit Function
            End If
        Next
    End If
    
End Function

