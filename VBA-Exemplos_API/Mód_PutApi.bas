Attribute VB_Name = "M�d_PutApi"
Option Explicit


Sub PutApi()

    Dim url     As String
    Dim req     As New MSXML2.ServerXMLHTTP60
    Dim body    As String
    
    Let url = "https://reqres.in/api/users/2"
    Let body = "{""name"":""Jo�o"",""job"":""Arquiteto""}"
    
        With req
            .Open "PUT", url, False
            .SetRequestHeader "Content-type", "application/json"
            .Send body
        End With
    
    Debug.Print req.Status
    Debug.Print req.StatusText
    Debug.Print req.ResponseText
    
    
End Sub


