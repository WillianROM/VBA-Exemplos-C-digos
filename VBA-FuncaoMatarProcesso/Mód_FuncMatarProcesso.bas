Attribute VB_Name = "M�d_FuncMatarProcesso"
Option Explicit

Function mataProcesso()

    Dim oServ As Object
    Dim cProc As Variant
    Dim oProc As Object
    
    
    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.ExecQuery("Select * from Win32_Process")
    
    For Each oProc In cProc
    
    'NOTA: � case sensitive
    
        If oProc.Name = "chromedriver.exe" Then
            oProc.Terminate
        End If
    
    Next

End Function
