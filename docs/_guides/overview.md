---
title: Overview
---

# Overview

VBA-Web consists of three primary components:

__WebClient__ executes requests and handles responses and is responsible for functionality shared between requests, such as authentication, proxy configuration, and security.
It is designed to be long-lived, maintaining state (e.g. authentication tokens) between requests.

__WebRequest__ is used to create detailed requests, including automatic formatting, querystrings, headers, cookies, and much more.

__WebResponse__ wraps http/cURL responses and includes automatically parsed data.

The following example is the standard form for a VBA-Web module:

1. A long-lived `WebClient` that sets the base url shared by all requests
2. A detailed `WebRequest` (that sets `Resource`, `Method`, and `Format`)
3. A wrapped response that includes automatically converted `Data` based on `Request.Format`

```vb
' Long-lived client, maintains state between requests
Private ClientInstance As WebClient
Public Property Get Client() As WebClient
    If ClientInstance Is Nothing Then
        ' Set base url shared by all requests
        Set ClientInstance = New WebClient
        ClientInstance.BaseUrl = "https://www.example.com/api/"
    End If

    Set Client = ClientInstance
End Property

Public Function GetProjects() As Collection
    ' Use GET and json (for request and response)
    Dim Request As New WebRequest
    Request.Resource = "projects"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.Json

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)

    ' -> GET https://www.example.com/api/projects
    '
    ' <- HTTP/1.1 200 OK
    '    ...
    '    {"data":[{"id":1,"name":"Project 1"},{"id":2,"name":"Project 2"}]}

    If Response.StatusCode <> WebStatusCode.Ok Then
        Err.Raise Response.StatusCode, "GetProjects", Response.Content
    Else
        ' Response is automatically converted to Dictionary/Collection by Request.Format
        Set GetProjects = Response.Data("data")
    End If
End Function
```
