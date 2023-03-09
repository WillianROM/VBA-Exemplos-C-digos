Attribute VB_Name = "Mód_01_verVersaoChromedriver"
Option Explicit


Sub VerificarVersaoChromedriver()
   ' Obter a versão do Chrome
    Dim versaoChrome As String
    versaoChrome = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version")
    
    ' Obter a versão do chromedriver
    Dim pathChromedriver As String
    pathChromedriver = Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic\chromedriver.exe"
    
    Dim versaoChromedriver As String
    versaoChromedriver = Split(CreateObject("WScript.Shell").Exec(pathChromedriver & " --version").StdOut.ReadAll, " ")(1)
    
    ' Comparar as versões
    If Split(versaoChrome, ".")(0) <> Split(versaoChromedriver, ".")(0) Then
        MsgBox "A versão " & Split(versaoChromedriver, ".")(0) & " do chromedriver na pasta " & Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic" & " não é compatível com a versão " & Split(versaoChrome, ".")(0) & " do Chrome instalado." & vbNewLine & _
        vbNewLine & "Por favor atualize o arquivo chromedriver na pasta mencionado anteriormente.", vbCritical, "Atualize o chromedriver"
        End
    End If
End Sub





