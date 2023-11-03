Attribute VB_Name = "MOD_Atualizar"
Option Explicit
' Utilize a blblioteca "Microsoft ActiveX Data Objects 6.1 Library"

Public Sub AtualizarDadosTabela()
    Dim rs              As ADODB.Recordset
    Dim conn            As ADODB.Connection
    Dim SQL             As String

    Dim TABELA          As ListObject
    Dim LINHA           As ListRow
    
    ' Conexão é o caminho até a lista
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    With conn
        ' https://www.connectionstrings.com/
        ' Para pegar a list, vai na engrenagem na página do Sharepoint -> Configuração da lista, daí peque na URL da página
        ' Desconsidere o %7B no inicio e o %7D no final que são as chaves em código
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes;" & _
                            "DATABASE=" & wsSharepoint.Range("B1") & ";" & _
                            "LIST={" & wsSharepoint.Range("B2") & "};"
        .Open
    End With
    
    
    
    
    
    ' Loopar tabela e inserir dados
    Set TABELA = wsTabela.ListObjects(1)
    
    For Each LINHA In TABELA.ListRows
    
        ' SQL é o comando
        SQL = "SELECT * FROM [" & wsTabela.ListObjects(1).Name & "] WHERE ID=" & LINHA.Range(, 1).Value
        
        rs.Open SQL, conn, adOpenDynamic, adLockOptimistic 'Necessário deixar a conexão openDynamic

        rs.Fields("ID_ESPELHO") = LINHA.Range(, 1).Value

        rs.Update
        
        'fechar recordset
        rs.Close

    Next LINHA
       
    
        
    'fechar conexão
    conn.Close

End Sub



