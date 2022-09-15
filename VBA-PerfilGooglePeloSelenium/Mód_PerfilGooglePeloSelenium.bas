Attribute VB_Name = "Mód_PerfilGooglePeloSelenium"
Option Explicit

Dim driver As New ChromeDriver

Sub ComoUtilizarPerfilGooglePeloSeleniumVBA()

    'Buscar perfil google salvo (cookies)
    
    If Dir("C:\SeleniumProfile", vbDirectory) = vbNullString Then
        MkDir "C:\SeleniumProfile"
    End If
    
    
    driver.SetProfile "C:\SeleniumProfile", True
    
    
    'Iniciar o acesso a plataforma
    
    With driver
        .Start "chrome", "https://mail.google.com/mail/"
        .Window.Maximize
        .Get "https://mail.google.com/mail/"
    End With


End Sub

