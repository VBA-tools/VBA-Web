Private pGitLabClient As WebClient
Private pToken As String

Private Property Get Token() As String
    If pToken = "" Then
        If Credentials.Loaded Then
            pToken = Credentials.Values("GitLab")("token")
        Else
            pToken = InputBox("GitLab Token?")
        End If
    End If
    
    Token = pToken
End Property

Private Property Get GitLabClient() As WebClient
    If pGitLabClient Is Nothing Then
        Set pGitLabClient = New WebClient
        pGitLabClient.BaseUrl = "https://gitlab.com/api/v4"
        
        Dim Auth As New TokenAuthenticator
        Auth.Setup _
            Header:="PRIVATE-TOKEN", _
            value:=Token
        Set pGitLabClient.Authenticator = Auth
    End If
    
    Set GitLabClient = pGitLabClient
End Property

Function GetSample() As String
    'On Error GoTo ErrorHandler
    
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Request.Resource = "Invoice.svc/sample"
    
    ' Set the request format (Set {format} segment, content-types, and parse the response)
    Request.Format = WebFormat.Json

    ' (GET, POST, PUT, DELETE, PATCH)
    Request.Method = WebMethod.HttpGet
    
    Set Response = GitLabClient.Execute(Request)
    
    If Response.StatusCode <> WebStatusCode.Ok Then
        Button = MsgBox(Response.StatusDescription, vbAbortRetryIgnore + vbCritical, "Hata olustu")
    Else
        GetSample = WebHelpers.ConvertToJson(Response.Data, " ", 2)
        'MsgBox WebHelpers.ConvertToJson(Response.Data, " ", 2)
    End If
    
'ErrorHandler:
'    MsgBox "The following error occurred: " & Err.Description
    
End Function

Function GetIssiues(ProjectId As Long, Start As Integer) As Object
    'On Error GoTo ErrorHandler
    
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Request.Resource = "/issues?project=" & ProjectId & "&page=" & Start & "&per_page=100"
    
    ' Set the request format (Set {format} segment, content-types, and parse the response)
    Request.Format = WebFormat.Json

    ' (GET, POST, PUT, DELETE, PATCH)
    Request.Method = WebMethod.HttpGet
    
    Set Response = GitLabClient.Execute(Request)
    
    If Response.StatusCode <> WebStatusCode.Ok Then
        Button = MsgBox(Response.StatusDescription, vbAbortRetryIgnore + vbCritical, "Hata olustu")
    Else
        Set GetIssiues = Response.Data
        'GetIssiues = WebHelpers.ConvertToJson(Response.Data, " ", 2)
        'MsgBox WebHelpers.ConvertToJson(Response.Data, " ", 2)
    End If
    
'ErrorHandler:
'    MsgBox "The following error occurred: " & Err.Description
    
End Function

Function GetEvents() As Object
    'On Error GoTo ErrorHandler
    
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Request.Resource = "/events?project=6452557&target_type=issue&action=closed&per_page=100"
    
    ' Set the request format (Set {format} segment, content-types, and parse the response)
    Request.Format = WebFormat.Json

    ' (GET, POST, PUT, DELETE, PATCH)
    Request.Method = WebMethod.HttpGet
    
    Set Response = GitLabClient.Execute(Request)
    
    If Response.StatusCode <> WebStatusCode.Ok Then
        Button = MsgBox(Response.StatusDescription, vbAbortRetryIgnore + vbCritical, "Hata olustu")
    Else
        Set GetEvents = Response.Data
    End If
    
End Function

Function GetNotes(ProjectId As Long, IssueId As Integer) As Object
    'On Error GoTo ErrorHandler
    
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Request.Resource = "/projects/" & ProjectId & "/issues/" & IssueId & "/notes"
    
    ' Set the request format (Set {format} segment, content-types, and parse the response)
    Request.Format = WebFormat.Json

    ' (GET, POST, PUT, DELETE, PATCH)
    Request.Method = WebMethod.HttpGet
    
    Set Response = GitLabClient.Execute(Request)
    
    If Response.StatusCode = WebStatusCode.NotFound Then
        'GetNotes = Null
    ElseIf Response.StatusCode <> WebStatusCode.Ok Then
        Button = MsgBox(Response.StatusDescription, vbAbortRetryIgnore + vbCritical, "Hata olustu")
    Else
        Set GetNotes = Response.Data
    End If
    
End Function
