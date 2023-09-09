VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGerar 
   Caption         =   "UserForm1"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   OleObjectBlob   =   "frmGerar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub txtPasta_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    On Error GoTo TratarErro

    Dim fdlg As FileDialog
    
    Set fdlg = Application.FileDialog(msoFileDialogFolderPicker)
    
    If fdlg.Show = -1 Then
        txtPasta.Text = fdlg.SelectedItems(1)
    Else
        MsgBox "Não foi selecionada uma pasta."
    End If
    
Sair:
    Exit Sub
TratarErro:
    MsgBox "Houve um erro na aplicação: " & Err.Description & " - " & Err.Number
    GoTo Sair
    
End Sub


Private Sub CommandButton1_Click()
    On Error GoTo TratarErro
    
    Dim llArquivo       As Long
    Dim lstrCaminho     As String
    Dim lRegiao         As Range
    Dim lTexto          As String
    Dim lLinhas         As Long
    Dim lLinha          As Long
    Dim lColunas        As Long
    Dim lColuna         As Long
    
    
    lstrCaminho = txtPasta & "\" & txtNome.Text
    
    Set lRegiao = Sheets("Planilha1").Range(frmGerar.refIntervalo)
    
    lLinhas = lRegiao.Cells.Rows.Count
    lColunas = lRegiao.Columns.Count
    
    'Verificar se o arquivo existe
    If Dir(lstrCaminho) = "" Then
    
        'Cria endereço de memória e abre o arquivo txt
        llArquivo = FreeFile
        
        Open lstrCaminho For Output As #llArquivo
        
        'Insere os dados de cada célula
        For lLinha = 1 To lLinhas
            lTexto = ""
            
            For lColuna = 1 To lColunas
            
                If lColuna = 1 Then
                    lTexto = lRegiao.Cells(lLinha, lColuna).Value
                Else
                    lTexto = lTexto & ";" & lRegiao.Cells(lLinha, lColuna).Value
                End If
                
            Next lColuna
            
            Print #llArquivo, lTexto
            
        Next lLinha
        
    Else
        MsgBox "Arquivo já existe"
        GoTo Sair
    End If
    
    MsgBox "Arquivo criado em: " & lstrCaminho
    
Sair:
    Close #llArquivo
    Exit Sub
    
TratarErro:
    MsgBox "Houm um erro na aplicação: " & Err.Description & " - " & Err.Number
    GoTo Sair
End Sub

Private Sub UserForm_Click()

End Sub
