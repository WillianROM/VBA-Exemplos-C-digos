Attribute VB_Name = "Mod_Delete"
Option Explicit
' Utilize a blblioteca "Microsoft ActiveX Data Objects 6.1 Library"

Public Sub DeletarTabelaInteira()
    Dim rs              As ADODB.Recordset
    Dim conn            As ADODB.Connection
    Dim SQL             As String
    
    If MsgBox("Voc� tem certeza que quer deletar?", vbYesNo, "ATEN��O") <> vbYes Then
        Exit Sub
    End If

    
    ' Conex�o � o caminho at� a lista
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    With conn
        ' https://www.connectionstrings.com/
        ' Para pegar a list, vai na engrenagem na p�gina do Sharepoint -> Configura��o da lista, da� peque na URL da p�gina
        ' Desconsidere o %7B no inicio e o %7D no final que s�o as chaves em c�digo
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes;" & _
                            "DATABASE=" & wsSharepoint.Range("B1") & ";" & _
                            "LIST={" & wsSharepoint.Range("B2") & "};"
        .Open
    End With
    
    
    ' SQL � o comando
    SQL = "DELETE * FROM [" & wsTabela.ListObjects(1).Name & "]"
    
    rs.Open SQL, conn

    
    'fechar recordset
    rs.Close
    
    'fechar conex�o
    conn.Close

End Sub


