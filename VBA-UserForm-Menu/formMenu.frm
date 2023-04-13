VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMenu 
   Caption         =   "Teste Menu Hambuguer"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490.001
   OleObjectBlob   =   "formMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'https://www.youtube.com/watch?v=JhzRIxq8G5Y&list=PLmOO8X35BgB3N8VYMgu9yBSz-Gg7RBFEV&index=3&ab_channel=RonanVico

Private Sub UserForm_Initialize()
    With FrameMenu
        .Top = 0 ' topo = 0 para colar o menu no topo
        .Left = -9999 ' left = -9999 para simplesmentes ocultar o frame
        .Height = formMenu.Height 'Deixa a altura do menu igual a form
    End With
End Sub

Private Sub btnMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoverMenu True
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If FrameMenu.Left = 0 Then
        MoverMenu False
    End If
End Sub

Private Function MoverMenu(Mostrar As Boolean)
    Dim leftFinal       As Long
    Dim leftInicial     As Long
    Dim cont            As Long
    Dim stepAFazer      As Long
    
    If Mostrar Then
        leftInicial = FrameMenu.Width * -1
        leftFinal = 0
        stepAFazer = 1
    Else
        leftInicial = 0
        leftFinal = FrameMenu.Width * -1
        stepAFazer = -1
    End If

    For cont = leftInicial To leftFinal Step stepAFazer
        FrameMenu.Left = cont
    Next cont

End Function
