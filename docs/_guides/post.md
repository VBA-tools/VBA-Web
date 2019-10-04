---
title: POST Request
draft: true
---

# POST Request

```vb
' (Use Client from Overview)

Public Function CreateProject(Project As Dictionary) As Long
    Dim Request As New WebRequest
    Request.Resource = "projects"
    Request.Method = WebMethod.HttpPost
    Request.Format = WebFormat.Json

    ' Body is automatically converted to json from Dictionary/Collection
    ' (based on Format or RequestFormat
    Set Response.Body = Project

    Dim Response As WebResponse
    Set Response = Client.Execute(Request)

    ' -> POST https://www.example.com/api/projects
    '
    '    {"name":"New Project"}
    '
    ' <- HTTP/1.1 201 Created
    '    ...
    '    {"data":{"id":3,"name":"new Project"}}

    If Response.StatusCode <> WebStatusCode.Created Then
        Err.Raise Response.StatusCode, "CreateProject", _
            "Failed to create project: " & Response.Content
    Else
        ' Return id of created project
        CreateProject = Response.Data("data")("id")
    End If
End Sub
```
