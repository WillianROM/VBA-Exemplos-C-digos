Attribute VB_Name = "Mod_Select"
Option Explicit
' Utilize a blblioteca "Microsoft ActiveX Data Objects 6.1 Library"

Public Sub SelectTabelaInteira()
    Dim rs              As ADODB.Recordset
    Dim conn            As ADODB.Connection
    Dim SQL             As String

    
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
    SQL = "SELECT * FROM [" & wsTabela.ListObjects(1).Name & "]"
    
    rs.Open SQL, conn
    
    ' Preencher os dados
    Dim FD          As ADODB.Field
    Dim col         As Long
    
    col = 1
    
    If rs.EOF = False Then

        For Each FD In rs.Fields
    
            With wsTabela.Cells(1, col)
                .Value = FD.Name
            End With
    
        col = col + 1
    
        Next FD
        
        'inserir dados do recordset na planilha
        wsTabela.Cells(2, 1).CopyFromRecordset rs
    
    End If
    
    
    
    'fechar recordset
    rs.Close
    
    'fechar conex�o
    conn.Close

End Sub
