Attribute VB_Name = "Mód_ConfigDownloadPDFNoChrome"
Option Explicit

Dim driver As New ChromeDriver

Sub Configurar_Download()
    
    If Dir("C:\SeleniumProfile", vbDirectory) = vbNullString Then
        MkDir "C:\SeleniumProfile"
    End If
    
    
    driver.SetProfile "C:\SeleniumProfile", True


    With driver
        .Start
        .Get ("chrome://settings/content/pdfDocuments")
    End With

    MsgBox "Antes de clicar em OK, altere para Download Automático na janela do Chrome que abriu", vbInformation, "Download Automático"

    driver.Quit
    
End Sub
