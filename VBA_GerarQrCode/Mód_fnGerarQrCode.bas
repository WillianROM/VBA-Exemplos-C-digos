Attribute VB_Name = "Mód_fnGerarQrCode"
Option Explicit

  Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
            Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
            ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function gerarQrCode(ByVal url As String, ByVal arquivo As String)
    
    Let url = "https://chart.googleapis.com/chart?cht=qr&chs=500x500&chl=" & url

    URLDownloadToFile 0, url, "C:\temp\" & arquivo & ".png", 0, 0
End Function
