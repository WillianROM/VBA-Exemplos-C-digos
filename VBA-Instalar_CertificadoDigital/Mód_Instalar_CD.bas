Attribute VB_Name = "Mód_Instalar_CD"

#If VBA7 Then
    'Para computadores de 64 bits
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long


    Public Declare PtrSafe Function FindWindowExW Lib "user32.dll" ( _
    ByVal hWnd1 As Long, _
    Optional ByVal hWnd2 As Long, _
    Optional ByVal lpsz1 As String, _
    Optional ByVal lpsz2 As String) As Long


    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal Hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long


    Public Declare PtrSafe Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal Hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long
    
#Else
'Para computadores de 32 bits
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long


    Public Declare Function FindWindowExW Lib "user32.dll" ( _
    ByVal hWnd1 As Long, _
    Optional ByVal hWnd2 As Long, _
    Optional ByVal lpsz1 As String, _
    Optional ByVal lpsz2 As String) As Long


    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long


    Public Declare Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

#End If

'=====================================================================

Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
ByVal Hwnd As Long, _
ByVal lpString As String, _
ByVal cch As Long) As Long



Public Function IdBotão(JanelaMãe, Texto)

    Dim TextoAPI As String * 255
inicio:
    
    ElementoFilho = FindWindowExW(JanelaMãe)
    ElementoFilho = FindWindowExW(ElementoFilho)
    ElementoFilho = FindWindowExW(ElementoFilho)
    
    
    Do While ElementoFilho <> 0
        TextoElemento = Left$(TextoAPI, GetWindowText(ElementoFilho, ByVal TextoAPI, 255))
        
        If TextoElemento Like "*" & Texto & "*" Then
           IdBotão = ElementoFilho
           Exit Function
        End If
      ElementoFilho = FindWindowExW(JanelaMãe, ElementoFilho)
    Loop


End Function

Sub ManipulandoJanela(ByVal senha_CD As String)

    'COMANDOS BÁSICOS
    CLICAR = &HF5
    FECHAR = &H10
    WM_SETTEXT = &HC
    
    'LOCALIZA A JANELA PELO TITULO
    Janela = FindWindow(vbNullString, "Assistente para Importação de Certificados") 'Assistente para Importação de Certificados
    
    'LOCALIZA BOTÃO PELO TEXTO
    Botão = IdBotão(Janela, "A&vançar")
    
    'ENVIA O COMANDO P/ O BOTÃO
    SendMessage Botão, CLICAR, 0, 0
    SendMessage Botão, CLICAR, 0, 0

'=============================================================================================
    Rem COLOCAR A SENHA
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Dim TextoAPI As String * 255
    
    Janela = FindWindow("NativeHWNDHost", "Assistente para Importação de Certificados")
    ElementoFilho = FindWindowExW(Janela)
    
    
    prevChild = 0
    currChild = 0
    
    WM_SETTEXT = &HC
    
    'Achar o Último filho

    Do
        currChild = FindWindowExW(ElementoFilho, prevChild, vbNullString, vbNullString)
        
        If currChild = 0 Then GoTo ulimo_neto
        
        TextoElemento = Left$(TextoAPI, GetWindowText(FindWindowExW(currChild), ByVal TextoAPI, 255))
        
        prevChild = currChild
    
    Loop While (currChild <> 0)



    'Unico filho
ulimo_neto:
    
    
    ElementoFilho = prevChild
    prevChild = 0
    
    Do
        currChild = FindWindowExW(ElementoFilho, prevChild, vbNullString, vbNullString)
        
        If currChild = 0 Then GoTo Achar_Campo_Senha
        
        
        TextoElemento = Left$(TextoAPI, GetWindowText(FindWindowExW(currChild), ByVal TextoAPI, 255))
        
        prevChild = currChild
    Loop While (currChild <> 0)



    'Achar o campo para colocar a senha
Achar_Campo_Senha:
    
    ElementoFilho = prevChild
    prevChild = 0
    
    Do
        currChild = FindWindowExW(ElementoFilho, prevChild, vbNullString, vbNullString)
        
        
        TextoElemento = Left$(TextoAPI, GetWindowText(currChild, ByVal TextoAPI, 255))
        
        If TextoElemento = "&Senha:" Then
            prevChild = currChild
            SendMessageByString FindWindowExW(ElementoFilho, prevChild, vbNullString, vbNullString), WM_SETTEXT, 0, senha_CD
        End If
        
        prevChild = currChild
    Loop While (currChild <> 0)


    '=============================================================================================
    'ENVIA O COMANDO P/ O BOTÃO
    For i = 1 To 6
        SendMessage Botão, CLICAR, 0, 0
    Next i


    '=============================================================================================
    'Clicar em Ok na caixa de mensagem
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Janela = FindWindow(vbNullString, "Assistente para Importação de Certificados")
    ElementoFilho = FindWindowExW(Janela)
    
    For i = 1 To 2
        SendMessage ElementoFilho, CLICAR, 0, 0
    Next i




End Sub


