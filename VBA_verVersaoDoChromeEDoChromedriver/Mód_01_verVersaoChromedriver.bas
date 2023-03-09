Attribute VB_Name = "M�d_01_verVersaoChromedriver"
Option Explicit


Sub VerificarVersaoChromedriver()
   ' Obter a vers�o do Chrome
    Dim versaoChrome As String
    versaoChrome = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version")
    
    ' Obter a vers�o do chromedriver
    Dim pathChromedriver As String
    pathChromedriver = Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic\chromedriver.exe"
    
    Dim versaoChromedriver As String
    versaoChromedriver = Split(CreateObject("WScript.Shell").Exec(pathChromedriver & " --version").StdOut.ReadAll, " ")(1)
    
    ' Comparar as vers�es
    If Split(versaoChrome, ".")(0) <> Split(versaoChromedriver, ".")(0) Then
        MsgBox "A vers�o " & Split(versaoChromedriver, ".")(0) & " do chromedriver na pasta " & Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic" & " n�o � compat�vel com a vers�o " & Split(versaoChrome, ".")(0) & " do Chrome instalado." & vbNewLine & _
        vbNewLine & "Por favor atualize o arquivo chromedriver na pasta mencionado anteriormente.", vbCritical, "Atualize o chromedriver"
        End
    End If
End Sub





