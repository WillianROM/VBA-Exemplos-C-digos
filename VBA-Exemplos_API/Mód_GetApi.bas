Attribute VB_Name = "Mód_GETApi"
Option Explicit

Sub GetApi()

    Dim url     As String
    Dim req     As New MSXML2.ServerXMLHTTP60
    
    
    Let url = "https://reqres.in/api/users?page=2"
    
        With req
            .Open "GET", url, False
            .SetRequestHeader "Content-type", "application/json"
            .Send
        End With
    
    Debug.Print req.Status
    Debug.Print req.StatusText
    Debug.Print req.ResponseText
    
    
End Sub





