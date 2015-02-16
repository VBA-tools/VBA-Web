---
title: Trello Analytics
---
<article class="content">

## Background

Engineering Inc. has set up a Kanban board in Trello for tracking an important business process. This has been a huge improvement over there existing email workflow and now they are looking to make the process even better with analytics, slow project tracking, and easy project creation.

## Goals

1. Load projects from Trello, including stage information, for analytics
2. Mark slow projects in Trello for prioritization
3. Create new projects

## Setup

An Excel Workbook has already been set up with analytics functionality and new project form, we just need to integrate Trello. The following functions need to be completed:

```VB.net
Public Function LoadProjects() As Collection
  ' TODO
End Function

Public Sub MarkSlowProjects(Projects As Collection)
  ' TODO
End Sub

Public Function CreateNewProject(ByRef Project)
  ' TODO
End Function
```

## Accessing Trello's API

Oftentimes, the most difficult part of working with APIs is getting everything set up to connect to them. Each one has its own set of required keys, tokens, and other items that need to be generated/retrieved in order to access the system.

From Trello's <a href="https://trello.com/docs/gettingstarted/" target="_blank">Getting Started</a> documentation:

1. Every request must contain an Application Key
2. To access private items, a User Token is required

Generating an Application Key for your account if fairly simple:

1. Log in to Trello
2. Visit <a href="https://trello.com/1/appKey/generate" target="_blank">https://trello.com/1/appKey/generate</a> to generate an Application Key (and Secret).

For this example, the Application Key is `595d...`.

Getting a User Token requires a little more work. For this example, a specific user (VBA-Web Bot) will be used to access Trello so a user token only needs to generated once. To generate a User Token:

1. Use the following URL: <br>`https://trello.com/1/authorize`<br>`?key=YOURAPPKEY&name=APP+NAME&expiration=EXPIRES`<br>`&response_type=token&scope=read,write`
2. Replace `YOURAPPKEY` with the Application Key you retrieved earlier (e.g. `595d...`)
3. Replace `APP+NAME` with the application's name (e.g. `Engineering+Inc`, with `+` for spaces)
4. Replace `EXPIRES` with `1day`, `30days`, or `never`
5. Go to the created URL in your browser, allow the application to use your account, and store the User Token for later.

For this example, the retrieved User Token is `ea3c...`.

## GET Board

With the Application Id and User Token ready, let's see if we can retrieve our Kanban board.

First, add a temporary `Test` method to the `Trello` module. The `BoardId` value can be found from the board URL when viewing the board (e.g. for `https://trello.com/b/iBpxWmUu/engineering-inc` the `BoardId` is `iBpxWmUu`).

```VB.net
' Trello.bas
Private Const ApplicationKey As String = "595d..."
Private Const UserToken As String = "ea3c..."
Private Const BoardId As String = "iBpxWmUu"

Sub Test
  ' TODO
End Sub
```

VBA-Web consists of three primary components: `WebClient`, `WebRequest`, and `WebResponse`.

### WebClient

`WebClient` executes requests and handles responses and is responsible for state/functionality that is shared between requests, such as authentication, proxy configuration, and security. For the Trello API, all requests start with "https://api.trello.com/1/" so this will be shared between all requests with `BaseUrl`.

```VB.net
Dim Client As New WebClient
Client.BaseUrl = "https://api.trello.com/1/"
```

### WebRequest

`WebRequest` is used to create detailed requests (including formatting, querystrings, headers, cookies, and much more). VBA-Web aims to make every part of the request configurable, so there are helpers to avoid building strings for URLs or Body values by hand and other tedious and potentially error-prone methods of creating requests.

```VB.net
Dim Request As New WebRequest

' Anti-pattern: Building URL by hand
Request.Resource = "boards/" & BoardId & "?key=" & ApplicationKey & "&token=" & UserToken

' Preferred
Request.Resource = "boards/{board_id}"
Request.AddUrlSegment "board_id", BoardId
Request.AddQuerystringParam "key", ApplicationKey
Request.AddQuerystringParam "token", UserToken

' Defaults:
' Request.Format = WebFormat.Json
' Request.Method = WebMethod.HttpGet
```

<h3>WebResponse</h3>

`WebResponse` wraps http and cURL repsonses and includes parsed `Data` based on `WebRequest.ResponseFormat`.

```VB.net
Dim Response As WebResponse
Set Response = Client.Execute(Request)

Debug.Print Response.StatusCode & ": " & Response.Content
```

All together:

```VB.net
' Trello.bas
Private Const ApplicationKey As String = "595d..."
Private Const UserToken As String = "ea3c..."
Private Const BoardId As String = "iBpxWmUu"

Sub Test
  Dim Client As New WebClient
  Client.BaseUrl = "https://api.trello.com/1/"

  Dim Request As New WebRequest
  Request.Resource = "boards/{board_id}"
  Request.AddUrlSegment "board_id", BoardId
  Request.AddQuerystringParam "key", ApplicationKey
  Request.AddQuerystringParam "token", UserToken

  Dim Response As WebResponse
  Set Response = Client.Execute(Request)

  Debug.Print Response.StatusCode & ": " & Response.Content
End Sub
```

## Debugging

Hopefully, the above test went smoothly, but if there were issues, how do you debug what happened?

### EnableLogging

Enable logging with `WebHelpers.EnableLogging = True` and open the Immediate Window (`View > Immediate Window` or `ctrl+g`) to view the raw request that was sent and response recieved.

```VB.net
Sub Test
  WebHelpers.EnableLogging = True
  ' ...
End Sub

' --> Request - #:##:## AM
' GET https://api.trello.com/1/boards/iBpxWmUu?key=595d...&token=ea3c...
' ...
'
' <-- Response - #:##:## AM
' 200 OK
' ...
'
' {"id":"5431d8cf70be14fc345c8e35","name":"Engineering Inc.",...}
```

</article>
