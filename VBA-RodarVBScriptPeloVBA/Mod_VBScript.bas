Attribute VB_Name = "Mod_VBScript"
Option Explicit

'Rodar arquivo VBScript pelo VBA

Dim Caminho     As String

Sub main()
    Caminho = ThisWorkbook.Path
    Call teste1
    Call Teste2
    Call Teste3
End Sub

'1. Pelo Prompt de Comando
Public Sub teste1()
    Shell "cmd.exe /c " & Caminho & "\VBScriptTeste.vbs"""
End Sub


'2. Usando o Windows Script
Public Sub Teste2()
    CreateObject("WScript.Shell").Run (Caminho & "\VBScriptTeste.vbs")
End Sub

Public Sub Teste3()
    Shell "WScript " & Caminho & "\VBScriptTeste.vbs"
End Sub
