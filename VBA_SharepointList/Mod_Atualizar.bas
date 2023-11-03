Attribute VB_Name = "MOD_Atualizar"
Option Explicit
' Utilize a blblioteca "Microsoft ActiveX Data Objects 6.1 Library"

Public Sub AtualizarDadosTabela()
    Dim rs              As ADODB.Recordset
    Dim conn            As ADODB.Connection
    Dim SQL             As String

    Dim TABELA          As ListObject
    Dim LINHA           As ListRow
    
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
    
    
    
    
    
    ' Loopar tabela e inserir dados
    Set TABELA = wsTabela.ListObjects(1)
    
    For Each LINHA In TABELA.ListRows
    
        ' SQL � o comando
        SQL = "SELECT * FROM [" & wsTabela.ListObjects(1).Name & "] WHERE ID=" & LINHA.Range(, 1).Value
        
        rs.Open SQL, conn, adOpenDynamic, adLockOptimistic 'Necess�rio deixar a conex�o openDynamic

        rs.Fields("ID_ESPELHO") = LINHA.Range(, 1).Value

        rs.Update
        
        'fechar recordset
        rs.Close

    Next LINHA
       
    
        
    'fechar conex�o
    conn.Close

End Sub



