Attribute VB_Name = "Mód_DeleteApi"
Option Explicit


Sub DeleteApi()

    Dim url     As String
    Dim req     As New MSXML2.ServerXMLHTTP60
    
    
    Let url = "https://reqres.in/api/users/2"

    
        With req
            .Open "PUT", url, False
            .SetRequestHeader "Content-type", "application/json"
            .Send
        End With
    
    Debug.Print req.Status
    Debug.Print req.StatusText
    Debug.Print req.ResponseText
    
End Sub

