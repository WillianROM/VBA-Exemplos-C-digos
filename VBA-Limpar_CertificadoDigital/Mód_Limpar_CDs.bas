Attribute VB_Name = "Mód_Limpar_CDs"
Option Explicit

Sub Limpar_Certificados_Digitais_do_Computador()

Dim strPath As String
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

strPath = "C:\Users\" & UsuarioRede & "\AppData\Roaming\Microsoft\SystemCertificates\My\Certificates"

 
'Delete all files in specified folder
fso.DeleteFile strPath & "\*.*"
 

'Excluir todos os arquivox pfx da pasta Downloads
strPath = Environ("userprofile") & "\Downloads"

On Error Resume Next
fso.DeleteFile strPath & "\*.pfx"


End Sub


Function UsuarioRede() As String
    Dim GetUserN
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    GetUserN = ObjNetwork.UserName
    UsuarioRede = GetUserN
End Function


