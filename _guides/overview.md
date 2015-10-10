---
title: Overview
---

# WebClient

`WebClient` executes requests and handles responses and is responsible for functionality shared between requests, such as authentication, proxy configuration, and security.

```VB.net
Dim Client As New WebClient
Client.BaseUrl = "https://www.example.com/api/"

Dim Auth As New HttpBasicAuthenticator
Auth.Setup Username, Password
Set Client.Authenticator = Auth

Dim Request As New WebRequest
Dim Response As WebResponse
' Setup WebRequest...

Set Response = Client.Execute(Request)
' -> Uses Http Basic authentication and appends Request.Resource to BaseUrl
```

# WebRequest

`WebRequest` is used to create detailed requests (including formatting, querystrings, headers, cookies, and much more).

```VB.net
Dim Request As New WebRequest
Request.Resource = "users/{Id}"

Request.Method = WebMethod.HttpPut
Request.RequestFormat = WebFormat.UrlEncoded
Request.ResponseFormat = WebFormat.Json

Dim Body As New Dictionary
Body.Add "name", "Tim"
Body.Add "project", "VBA-Web"
Set Request.Body = Body

Request.AddUrlSegment "Id", 123
Request.AddQuerystringParam "api_key", "abcd"
Request.AddHeader "Authorization", "Token ..."

' -> PUT (Client.BaseUrl)users/123?api_key=abcd
'    Authorization: Token ...
'
'    name=Tim&project=VBA-Web
```

# WebResponse

`WebResponse` wraps http/cURL responses and includes parsed `Data` based on `Request.ResponseFormat`.

```VB.net
Dim Response As WebResponse
Set Response = Client.Execute(Request)

If Response.StatusCode = WebStatusCode.Ok Then
  ' Response.Headers, Response.Cookies
  ' Response.Data -> Parsed Response.Content based on Request.ResponseFormat
  ' Response.Body -> Raw response bytes
Else
  Debug.Print "Error: " & Response.StatusCode & " - " & Response.Content
End If
```