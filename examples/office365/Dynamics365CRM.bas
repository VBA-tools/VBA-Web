Public Function ObjectRequest(ObjectName As String, Top As Integer, Filter As String) As WebRequest
    Dim Request As New WebRequest
    Request.Resource = "{ObjectName}?$select=name&$top={Top}&$filter=contains(name,'{Filter}')"
    Request.AddUrlSegment "ObjectName", ObjectName
    Request.AddUrlSegment "Top", Top
    Request.AddUrlSegment "Filter", Filter
    Request.ResponseFormat = json
    Request.Method = HttpGet
    Set ObjectRequest = Request
End Function

Public Function FindDynamicsCrmAccount(Filter As String) as String
    Dim Client As New WebClient
    Dim Response As WebResponse
    Dim RestUrl As String
    Dim ApiVersion As String
    Dim Data As Object

    RestUrl = "https://<your_company>.crm4.dynamics.com" ' crm4 is europe, lookup you own api url
    ApiVersion = "9.0"
    ' Create a new IAuthenticator implementation
    Dim Auth As New Office365Authenticator
    ' Setup authenticator...
    Auth.Setup "2ad88395-b77d-4561-9441-d0e40824f9bc", RestUrl, "<your_user>", "<your_password>"
    Auth.TokenUrl = "https://login.microsoftonline.com/common/oauth2/token"
    ' Attach the authenticator to the client
    Set Client.Authenticator = Auth
    Client.BaseUrl = RestUrl & "/api/data/v" & ApiVersion & "/"

    Set Response = Client.Execute(ObjectRequest("accounts", 1, Filter)) 'search for accounts, using filter, return 1 result
    If (Response.StatusCode = Ok) Then
        FindDynamicsCrmAccount = Response.Data("value")(1)("name")
    Else
        MsgBox "Connection with Dynamics CRM failed", vbCritical, "Error"
    End If
End Function