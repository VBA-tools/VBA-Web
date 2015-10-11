---
title: Overview
---

VBA-Web consists of three primary components:

__WebClient__ executes requests and handles responses and is responsible for functionality shared between requests, such as authentication, proxy configuration, and security.
It is designed to be long-lived, maintaining state (e.g. authentication tokens) between requests.

__WebRequest__ is used to create detailed requests, including automatic formatting, querystrings, headers, cookies, and much more.

__WebResponse__ wraps http/cURL responses and includes automatically parsed data.

The following example is the standard form for a VBA-Web module:

1. A long-lived `WebClient` that maintains authentication state between requests. The example uses HTTP Basic authentication with the `HttpBasicAuthenticator`, which can be installed with the VBA-Web installer or from the `authenticators` folder
2. A detailed `WebRequest` that sets `Method` and `Format` (or `RequestFormat` and `ResponseFormat` if different formats are needed) and uses url-formatting with `AddUrlSegment` and automatic body conversion from `Dictionary` or `Collection` based on the set request format
3. The response includes the `StatusCode`, `Content` (raw response string), `Data` (converted response based on response format), and status description, body (bytes), headers, and cookies

```VB.net
' Long-lived client, maintains state between requests
Private ClientInstance As WebClient
Public Property Get Client() As WebClient
    If ClientInstance Is Nothing Then
        ' Set base url shared by all requests and use HTTP Basic authentication
        Set ClientInstance = New WebClient
        ClientInstance.BaseUrl = "https://www.example.com/api/"

        Dim Auth As New HttpBasicAuthenticator
        Auth.Setup "username", "password"
        Set ClientInstance.Authenticator = Auth
    End If

    Set Client = ClientInstance
End Property

Public Sub UpdateProject(Project As Dictionary)
    ' Use PUT, format url, and automatically convert Dictionary to json
    Dim Request As New WebRequest
    Request.Resource = "projects/{id}"
    Request.Method = WebMethod.HttpPut
    Request.Format = WebFormat.Json

    Request.AddUrlSegment "id", Project("id")
    Set Request.Body = Project

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)

    ' -> PUT https://www.example.com/api/projects/123
    '    Authorization: Basic ...(set from Authenticator)
    '
    '    {"id":123,"name":"Project Name"}

    ' <- HTTP/1.1 204 No Content

    If Response.StatusCode <> WebStatus.NoContent Then
        Err.Raise Response.StatusCode, "UpdateProject", Response.Content
    End If
End Function
```
