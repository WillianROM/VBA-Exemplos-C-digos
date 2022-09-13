Attribute VB_Name = "MóduloSQL"
Option Explicit
Sub main()

    Call CopiarListaSQL_BookOficial("BASE")

End Sub


Function ConexaoBD_Book_Oficial()
    Dim arq As String
    
    arq = ThisWorkbook.Path & "\BaseDados.accdb"
    
    ConexaoBD_Book_Oficial = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & arq & ";Persist Security Info=False"

End Function

Function RetornaSQL(vQuery As String)

    Select Case vQuery
    
        Case 1
            RetornaSQL = "SELECT * FROM Base"
    
    End Select

End Function
Sub CopiarListaSQL_BookOficial(Planilha As String)
    Dim vUltCelSel As Range
    
    Set vUltCelSel = ActiveCell
    
    Dim col     As Integer
    Dim i       As Long
    Dim SQL     As String
    Dim arq     As String
    Dim cm      As New ADODB.Connection
    Dim rs      As New ADODB.Recordset
    Dim FD      As ADODB.Field
    Dim vcol    As Range
    Dim vRng    As Range
    Dim w       As Worksheet
    
    Set w = Worksheets(Planilha)


    'Criar conexão com o BD
    Set cm = New ADODB.Connection
    
    'abrir conexão
    cm.Open ConexaoBD_Book_Oficial
    
    'Criar um recordset
    Set rs = New ADODB.Recordset
    
    SQL = RetornaSQL(1)
    
    'Realiza a Consulta
    
    rs.Open SQL, cm

    'verifica se há dados no recordset
    col = 1
    
    If rs.EOF = False Then
    
        For Each FD In rs.Fields
    
            With w.Cells(1, col)
                .Value = FD.Name
            End With
    
        col = col + 1
    
        Next FD
        
    'inserir dados do recordset na planilha
        w.Cells(2, 1).CopyFromRecordset rs
    
    
    End If
    
    'fechar recordset
    rs.Close
    
    'fechar conexão
    cm.Close

End Sub
