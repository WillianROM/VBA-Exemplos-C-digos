Attribute VB_Name = "M�d_NavegarSemInstalarSelenium"
Option Explicit

'https://www.youtube.com/watch?v=NBKc06hpTow&list=TLPQMzEwMTIwMjM9GfdaEEQxtg&index=21&ab_channel=RonanVico

'Necess�rio:
    '* Fazer um m�dulo com os c�digos do github: https://github.com/VBA-tools/VBA-JSON/blob/master/JsonConverter.bas
    '1. No Editor de VBA, clique em "Ferramentas" -> "Refer�ncias".
    '2. Na janela "Refer�ncias - Projeto", selecione a op��o "Microsoft Scripting Runtime".
    '3. Clique em "OK".
    
'Aten��o:
    ' � necess�rio que o driver esteja aberto

Public Const PORT      As String = "9515"
Public Const url       As String = "http://localhost:" & PORT & "/"

Public Function SEND_REQUEST(url As String, Optional body As String, Optional METHOD As String = "GET")
    Dim HReq    As Object
    Dim resp    As String
    
    Set HReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With HReq
        .Open METHOD, url, False
        .send body
        resp = .responsetext
    End With
    
    SEND_REQUEST = resp
    
End Function

Sub AcessarGoogleSemSelenium()
Rem - Procure pela sess�o "Command Reference" na https://www.selenium.dev/documentation/legacy/json_wire_protocol/
    Dim body                    As String
    Dim rep                     As String
    Dim auxUrl                  As String
    Dim objRespostaNavegador
    Dim objRespostaElemento
    
Rem - Determinar qual navegador ser� utilizado, nesse caso � o chrome
    body = "{ ""desiredCapabilities"": { ""caps"": { ""nativeEvents"": false, ""browserName"": ""chrome"", ""version"": """", ""platform"": ""ANY"" } } }"
    rep = SEND_REQUEST(url & "session", body, "POST")
    
    Set objRespostaNavegador = JsonConverter.ParseJson(rep)

Rem - Entrar no site do Google
    'session/:sessionid/url
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/url"
    body = "{""url"":""https://www.google.com.br""}"
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
Rem - Pegar um elemento da p�gina - nesse caso usando xpath para pegar o campo input
    'session/:sessionid/element
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/element"
    body = "{""using"":""xpath"",""value"":""//input[@name='q']""}"
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
    Set objRespostaElemento = JsonConverter.ParseJson(rep)
    

Rem - Escrever algo no campo input pego anteriormente
    'session/:sessionid/element/:id/value
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/element/" & objRespostaElemento("value")("ELEMENT") & "/value"
    body = "{""value"":[""VBA Selenium Basic""]}"
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
 Rem - Pegar um elemento da p�gina - nesse caso usando xpath para pegar o campo bot�o
    'session/:sessionid/element
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/element"
    body = "{""using"":""name"",""value"":""btnK""}"
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
    Set objRespostaElemento = JsonConverter.ParseJson(rep)
    
Rem - Enviar um click no bot�o
    'session/:sessionid/element/:id/value
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/element/" & objRespostaElemento("value")("ELEMENT") & "/click"
    body = ""
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
End Sub



Sub AcessarGoogleSemSeleniumUsandoScriptDoJavaScript()
Rem - Procure pela sess�o "Command Reference" na https://www.selenium.dev/documentation/legacy/json_wire_protocol/
    Dim body                    As String
    Dim rep                     As String
    Dim auxUrl                  As String
    Dim objRespostaNavegador
    Dim objRespostaElemento
    
Rem - Determinar qual navegador ser� utilizado, nesse caso � o chrome
    body = "{ ""desiredCapabilities"": { ""caps"": { ""nativeEvents"": false, ""browserName"": ""chrome"", ""version"": """", ""platform"": ""ANY"" } } }"
    rep = SEND_REQUEST(url & "session", body, "POST")
    
    Set objRespostaNavegador = JsonConverter.ParseJson(rep)

Rem - Entrar no site do Google
    'session/:sessionid/url
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/url"
    body = "{""url"":""https://www.google.com.br""}"
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
    'ssssion/:sessionId/execute
    auxUrl = url & "session/" & objRespostaNavegador("sessionId") & "/execute"
    body = "{""script"":""document.getElementsByName('q').item(0).setAttribute('value','VBA Selenium Basic')"",""args"":[]}"
    
    rep = SEND_REQUEST(auxUrl, body, "POST")
    
End Sub
