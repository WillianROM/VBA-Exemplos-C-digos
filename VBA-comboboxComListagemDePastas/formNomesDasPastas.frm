VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNomesDasPastas 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formNomesDasPastas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formNomesDasPastas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim arr
    Dim fso     As Scripting.FileSystemObject
    Dim fd      As Folder
    Dim sFD     As Folder
    Dim Cont    As Long
    
    'Criar uma listagem de nome das pastas e colocar no combobox
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fd = fso.GetFolder(ThisWorkbook.Path)
    
    ReDim arr(1 To fd.SubFolders.Count, 1 To 1)
    
    For Each sFD In fd.SubFolders
    
        Cont = Cont + 1
        arr(Cont, 1) = sFD.Name

    Next sFD
    
    cboSubPastas.List = arr
    
End Sub


