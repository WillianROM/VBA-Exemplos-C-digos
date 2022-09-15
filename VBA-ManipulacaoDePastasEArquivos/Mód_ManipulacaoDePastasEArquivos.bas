Attribute VB_Name = "Mód_Certificados"
Option Explicit


Sub Salvar_certificados_em_outra_Pasta()

    Dim sfol As String, dfol As String
    
    sfol = Environ("userprofile") & "\AppData\Roaming\Microsoft\SystemCertificates\My\Certificates"
    dfol = "C:\certificados"
    
    Call CriaPasta(dfol)
    Call MoverTodosOsFicheiros(sfol, dfol)

End Sub

Sub Devolver_certificados_para_pasta_Certificates()

    Dim sfol As String, dfol As String
    
    sfol = "C:\certificados"
    dfol = Environ("userprofile") & "\AppData\Roaming\Microsoft\SystemCertificates\My\Certificates"
    
    Call MoverTodosOsFicheiros(sfol, dfol)
    Call ApagarPastaExistente(sfol)

End Sub



Sub CriaPasta(dfol As String)

    Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dfol) Then
        fso.CreateFolder (dfol)
    End If
    
End Sub



Sub MoverTodosOsFicheiros(sfol As String, dfol As String)

    Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
        fso.MoveFile (sfol & "\*.*"), dfol

End Sub


Sub ApagarPastaExistente(fol As String)

    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFolder fol

End Sub

