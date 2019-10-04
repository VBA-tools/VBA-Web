---
title: GET Request
---

# GET Request

The following demonstrates a simple GET request:

1. `Request.Format` sets both the `RequestFormat`, which sets the `Content-Type` header and automatically converts the request body, and `ResponseFormat`, which sets the `Accept` header and automatically converts the response body
2. `Request.Method` uses the `WebMethod` enum to set the request method (`GET`, `POST`, `PUT`, `PATCH`, and `HEAD` are available)
3. `Client.GetJson` is a shortcut for executing standard `GET` + `json` requests

```vb
' (Use Client from Overview)

Public Function GetProject(Id As Long) As Dictionary
    Dim Request As New WebRequest
    Request.Resource = "projects/" & Id

    ' Set request and response format
    ' - sets Content-Type and Accept headers
    ' - converts request and response bodies
    Request.Format = WebFormat.Json

    ' Method: HttpGet = GET
    ' POST, PUT, DELETE, PATCH, HEAD also available
    Request.Method = WebMethod.HttpGet

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)

    ' -> GET https://www.example.com/api/projects/1
    '
    ' <- HTTP/1.1 200 OK
    '    ...
    '    {"data":{"id":1,"name":"Project 1"}}

    If Response.StatusCode = WebStatusCode.Ok Then
        ' json response is automatically parsed based Request.Format
        Set GetProject = Response.Data("data")
    End If
End Function

Public Function GetProject2(Id As Long) As Dictionary
    ' For GET + json, GetJson can be used
    ' (equivalent to GetProject above)
    Dim Response As WebResponse
    Set Response = Client.GetJson("projects/" & Id)

    If Response.StatusCode = WebStatusCode.Ok Then
        Set GetProject2 = Response.Data("data")
    End If
End Function
```
