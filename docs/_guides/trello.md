---
title: Trello
draft: true
---
## Background

Engineering Inc. has set up a Kanban board in Trello for tracking an important business process. This has been a huge improvement over there existing email workflow and now they are looking to make the process even better with analytics, slow project tracking, and easy project creation.

## Goals

1. Load projects from Trello, including stage information, for analytics
2. Mark slow projects in Trello for prioritization
3. Create new projects


## Setup

An Excel Workbook has already been set up with analytics functionality and new project form, we just need to integrate Trello. The following functions need to be completed:

```vb
Public Function LoadProjects() As Collection
  ' TODO
End Function

Public Sub MarkSlowProjects(Projects As Collection)
  ' TODO
End Sub

Public Sub CreateProject(ByRef Project)
  ' TODO
End Sub
```

## Accessing Trello's API

Oftentimes, the most difficult part of working with APIs is getting everything set up to connect to them. Each one has its own set of required keys, tokens, and other items that need to be generated/retrieved in order to access the system.

From Trello's <a href="https://trello.com/docs/gettingstarted/" target="_blank">Getting Started</a> documentation:

1. Every request must contain an Application Key
2. To access private items, a User Token is required

Following the directions in the <a href="https://trello.com/docs/gettingstarted/" target="_blank">Getting Started</a> documentation, an Application Key and User Token are created. A specific user (VBA-Web Bot) will be used to access Trello so a user token only needs to generated once, but your app may need to generate tokens for other users.

For this example, the following example values will be used:

- Application Key: `key...`
- User Token: `token...`

## GET Board

With the Application Id and User Token ready, let's run a quick test to see if we can retrieve our Kanban board.

Add a temporary `GetBoard` method to the `Trello` module. The `BoardId` value can be found from the board URL when viewing the board (e.g. for `https://trello.com/b/iBpxWmUu/engineering-inc` the `BoardId` is `iBpxWmUu`).

1. Create a `WebClient` that will handle requests and responses and is responsible for shared functionality like authentication, proxy configuration, and security. For the Trello API, all requests start with "https://api.trello.com/1/" so this will be shared between all requests with `BaseUrl`.
2. `WebRequest` is used to create detailed requests (including formatting, querystrings, headers, cookies, and much more). VBA-Web aims to make every part of the request configurable, so there are helpers to avoid building strings for URLs or Body values by hand and other tedious and potentially error-prone methods of creating requests.
3. `WebResponse` wraps http and cURL repsonses and includes parsed `Data` based on `WebRequest.ResponseFormat`.

```vb
' Trello.bas
Private Const ApplicationKey As String = "key..."
Private Const UserToken As String = "token..."
Private Const BoardId As String = "iBpxWmUu"

Sub GetBoard()
  Dim Client As New WebClient
  Client.BaseUrl = "https://api.trello.com/1/"

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

  Dim Response As WebResponse
  Set Response = Client.Execute(Request)

  Debug.Print Response.StatusCode & ": " & Response.Content
End Sub
```

## Debugging

Hopefully, the above test went smoothly, but if there were issues, how do you debug what happened?

Enable logging with `WebHelpers.EnableLogging = True` and open the Immediate Window (`View > Immediate Window` or `ctrl+g`) to view the raw request that was sent and response recieved.

```vb
Sub GetBoard()
  WebHelpers.EnableLogging = True
  ' ...
End Sub

' --> Request - #:##:## AM
' GET https://api.trello.com/1/boards/iBpxWmUu?key=key...&token=token...
' ...
'
' <-- Response - #:##:## AM
' 200 OK
' ...
'
' {"id":"5431d8cf70be14fc345c8e35","name":"Engineering Inc.",...}
```

## LoadProjects

With the `GetBoard` test successful, we're ready to start work on the `LoadProjects` method. First, let's check out what we need for the `KanbanProject` class.

```vb
...
```

Examining Trello's API docs, we can get this information with...

```vb
' Trello.bas
Private Const ApplicationKey As String = "key..."
Private Const UserToken As String = "token..."
Private Const BoardId As String = "iBpxWmUu"

Function LoadProjects() As Collection
  Dim Client As New WebClient
  Client.BaseUrl = "https://api.trello.com/1/"

  Dim Request As New WebRequest
  Request.Resource = "boards/{board_id}"
  Request.AddUrlSegment "board_id", BoardId
  Request.AddQuerystringParam "key", ApplicationKey
  Request.AddQuerystringParam "token", UserToken

  Dim Response As WebResponse
  Set Response = Client.Execute(Request)

  Dim Projects As New Collection
  Dim Project As KanbanProject
  Dim Card As Dictionary

  For Each Card In Response.Data("cards")
    Set Project = New KanbanProject

    ' TODO...

    Projects.Add Project
  Next Card

  Set LoadProjects = Projects
End Collection
```
