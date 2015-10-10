Attribute VB_Name = "Todoist"
Dim pClient As WebClient
Public Property Get TodoistClient() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "https://todoist.com/api/v6/"
        
        Dim Auth As New TodoistAuthenticator
        Auth.Setup CStr(Credentials.Values("Todoist")("id")), CStr(Credentials.Values("Todoist")("secret")), CStr(Credentials.Values("Todoist")("redirect_url"))
        Auth.Scope = "data:read"
        Auth.Login
        
        Set pClient.Authenticator = Auth
    End If
    
    Set TodoistClient = pClient
End Property

Public Sub LoadProjects()
    ' See https://developer.todoist.com/#retrieve-data
    Dim Request As New WebRequest
    Request.Resource = "sync"
    Request.AddQuerystringParam "seq_no", 0
    Request.AddQuerystringParam "seq_no_global", 0
    Request.AddQuerystringParam "resource_types", "[""projects""]"
    
    Dim Response As WebResponse
    Set Response = TodoistClient.Execute(Request)
    
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print "Project Count: " & Response.Data("Projects").Count
    End If
End Sub

