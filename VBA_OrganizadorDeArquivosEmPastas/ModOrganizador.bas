Attribute VB_Name = "ModOrganizador"
Option Explicit

' O código original está disponível em https://organicsheets.top/organizador-de-pastas/
' Esse código abaixo foi modificado por Willian Rafael de Oliveira Melo em 20/07/2024
' Para a execução, apenas coloque o arquivo com esse código na mesma pasta que deverá organizar os arquivos e executar

Sub organizador()
    

    Application.ScreenUpdating = False
    
    
    On Error GoTo ErrorHandler
        
    Dim fso             As Scripting.FileSystemObject
    Dim Pasta
    
    Dim caminho         As String
    Let caminho = GetWorkbookPath & "\"
    
    Dim arquivos        As Collection
    Dim arquivo         As Variant
    'Call Lista_Arquivos_nas_pastas
    
    Set arquivos = ListFilesInFolder(caminho, True)
    
    
    'Definindo as variáveis de cada pasta
    Dim musicas_dir     As String
    Let musicas_dir = caminho & "Musicas"
    
    Dim docs_dir        As String
    Let docs_dir = caminho & "Docs"
    
    Dim exec_dir        As String
    Let exec_dir = caminho & "Executáveis"
    
    Dim planilhas_dir   As String
    Let planilhas_dir = caminho & "Planilhas"
    
    Dim powerpoint_dir  As String
    Let powerpoint_dir = caminho & "Apresentações"
    
    Dim foto_dir        As String
    Let foto_dir = caminho & "Fotos"
    
    Dim video_dir       As String
    Let video_dir = caminho & "Vídeos"
    
    Dim outros_dir      As String
    Let outros_dir = caminho & "Outros"
    
    Dim compactados_dir As String
    Let compactados_dir = caminho & "Compactados"
    
    'Usando o Dir para conferir se a pasta existe, caso não exista usa o MkDir para criação da pasta
    If Dir(musicas_dir, vbDirectory) = "" Then
        MkDir musicas_dir
    End If
    
    If Dir(docs_dir, vbDirectory) = "" Then
        MkDir docs_dir
    End If
    
    If Dir(exec_dir, vbDirectory) = "" Then
        MkDir exec_dir
    End If
    
    If Dir(planilhas_dir, vbDirectory) = "" Then
        MkDir planilhas_dir
    End If
    
    If Dir(powerpoint_dir, vbDirectory) = "" Then
        MkDir powerpoint_dir
    End If
    
    If Dir(foto_dir, vbDirectory) = "" Then
        MkDir foto_dir
    End If
    
    If Dir(video_dir, vbDirectory) = "" Then
        MkDir video_dir
    End If
    
    If Dir(outros_dir, vbDirectory) = "" Then
        MkDir outros_dir
    End If
    
    If Dir(compactados_dir, vbDirectory) = "" Then
        MkDir compactados_dir
    End If
    
    
    'Definindo as variáveis com as arrays com os formatos de cada tipo de arquivo
    Dim musicas()       As Variant
    Let musicas = Array("mp3", "wav")
    
    Dim docs()          As Variant
    Let docs = Array("doc", "pdf", "txt", "docx", "prn", "dif", "log")
    
    Dim exec()          As Variant
    Let exec = Array("exe")
    
    Dim planilhas()     As Variant
    Let planilhas = Array("xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls", "xlt", "xlam", "xml", "csv", "iqy")
    
    Dim powerpoint()    As Variant
    Let powerpoint = Array("ppt", "pps", "pptx")
    
    Dim foto()          As Variant
    Let foto = Array("jpeg", "jpg", "gif", "bmp", "wmf", "png", "ico")
    
    Dim video()         As Variant
    Let video = Array("vob", "mov", "mp4", "mpg", "avi")
    
    Dim compactados()   As Variant
    Let compactados = Array("zip", "rar", "arj", "cab", "tar")
  
        
    'loop por todos os itens já listados na planilha e movendo-os para as pastas correspondentes
    For Each arquivo In arquivos
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next
        
        If IsInArray(LCase(arquivo(2)), musicas) Then
            fso.MoveFile (caminho) & arquivo(1), (musicas_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), docs) Then
            fso.MoveFile (caminho) & arquivo(1), (docs_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), exec) Then
            fso.MoveFile (caminho) & arquivo(1), (exec_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), planilhas) Then
            fso.MoveFile (caminho) & arquivo(1), (planilhas_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), powerpoint) Then
            fso.MoveFile (caminho) & arquivo(1), (powerpoint_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), foto) Then
            fso.MoveFile (caminho) & arquivo(1), (foto_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), video) Then
            fso.MoveFile (caminho) & arquivo(1), (video_dir) & "\" & arquivo(1)
            
        ElseIf IsInArray(LCase(arquivo(2)), compactados) Then
            fso.MoveFile (caminho) & arquivo(1), (compactados_dir) & "\" & arquivo(1)
            
        Else
            fso.MoveFile (caminho) & arquivo(1), (outros_dir) & "\" & arquivo(1)
        End If
        
                
    Next arquivo

    
    Application.ScreenUpdating = True
    
    Principal.Select
    
    MsgBox "Itens já Organizados", vbInformation, "Organizador"
    
    Exit Sub
    
ErrorHandler:

    Principal.Select
    Application.ScreenUpdating = True
        
    MsgBox "Certifique se o caminho Do diretório está correto." & vbNewLine & Err.Description, vbCritical, "Erro: " & Err.Number
    
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    'funcao para identificar se algum item corresponde ao conjunto de array
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function ListFilesInFolder(SourceFolderName As String, IncludeSubfolders As Boolean) As Collection

    Dim fso             As Scripting.FileSystemObject
    Dim SourceFolder    As Scripting.Folder
    Dim FileItem        As Scripting.File
    Dim dados           As Collection
    Dim fileInfo        As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = fso.GetFolder(SourceFolderName)
    Set dados = New Collection ' Inicializa a coleção
    
    For Each FileItem In SourceFolder.Files
        If FileItem.Name <> ThisWorkbook.Name And FileItem.Name <> "~$" & ThisWorkbook.Name Then
            fileInfo = Array(FileItem.ParentFolder, FileItem.Name, fso.GetExtensionName(FileItem))
            dados.Add fileInfo ' Adiciona o array de informações à coleção
        End If
    Next FileItem
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set fso = Nothing
    
    Set ListFilesInFolder = dados ' Retorna a coleção
    
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
        Dim objRegistryProvider                 As Object
        Dim strRegistryPath                     As String
        Dim arrSubKeys()
        Dim strSubKey                           As Variant
        Dim strUrlNamespace                     As String
        Dim strMountPoint                       As String
        Dim strLocalPath                        As String
        Dim strRemainderPath                    As String
        Dim strLibraryType                      As String
    
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


