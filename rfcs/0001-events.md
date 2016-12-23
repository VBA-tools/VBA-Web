- Start Date: 2016-12-12
- RFC PR: https://github.com/VBA-tools/VBA-Web/pull/266
- VBA-Web Issue: (leave this empty)

# Summary

Currently, VBA-Web has a two mechanisms for extension: Authenticators and Custom Formatters. Authenticators have access to the http and curl objects when they're being prepared, the request before it's executed, and the response. Custom Formatters have access to the response body. These are very valuable extension points for adding more advanced functionality to VBA-Web and a more generalized system based on events would simplify these two existing extension types and allow for other customized functionality in the future.

# Motivation

The event system could provide the foundation for the two existing extension types and be used for more advanced functionality, all while simplifying the core code. `IWebAuthenticator` could be removed with Authenticators using events, more advanced Custom Formatters could be added for things other than just converting the request/response body, and it would be straightforward to add async functionality in the future.

# Detailed design

## Events

The request lifecycle for VBA-Web is as follows:

```txt
Client.Execute()
|
Setup http/curl
|
Execute request
|
Receive response
|
Return response
```

The following events can be added in between each stage:

```txt
Client.Execute()
|
Setup http/curl
| -> a) BeforeRequest
| -> b) PrepareHttp
| -> c) PrepareCurl
Execute request
| -> d) Execute
Receive response
| -> e) Parse
| -> f) Response
Return response
```

#### a) `Client_BeforeRequest(ByRef Request As WebRequest)`

This event is called before the request is generated from the `WebReqest`, so that changes can be made to `WebRequest` before it is parsed and executed (e.g. setting authentication headers, custom parsing, etc.).

#### b) `Client_PrepareHttp(Request As WebRequest, ByRef Http As WinHttpRequest)` (Windows)

Called after the `WinHttpRequest` has been prepared, this event can be used to interact with the `WinHttpRequest` directly (e.g. setting options, authentication, etc.).

#### c) `Client_PrepareCurl(Request As WebRequest, ByRef Curl As String)` (Mac)

Called after the curl command has been prepared (occurs after `BeforeRequest` since the request is parsed for the curl command), this event can be used to interact with the curl command directly (e.g. adding flags).

#### d) `Client_Execute(Request As WebRequest)`

This event occurs when the request is sent, before a response has been received.

#### e) `Client_Parse(Request As WebRequest, ByRef Response As WebResponse)`

This event occurs after the response has been initially parsed, but before it has been returned, and can be used to perform any custom parsing / response handling.

#### f) `Client_Response(Request As WebRequest, Response As WebResponse)`

This event occurs after the entire request lifecycle is complete and the response parsing is finalized.

## Usage

The event system is designed to be used primarily for extensions, but may be built off for future core functionality. Below is an example of converting the existing HttpBasicAuthenticator to events

Before:

```vb
' HttpBasicAuthenticator.cls
Implements IWebAuthenticator

Private Const web_HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0

Public Username As String
Public Password As String

Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    Request.SetHeader "Authorization", "Basic " & WebHelpers.Base64Encode(Me.Username & ":" & Me.Password)
End Sub

Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
End Sub

Private Sub IWebAuthenticator_PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
    Http.SetCredentials Me.Username, Me.Password, web_HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
End Sub

Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    Curl = Curl & " --basic --user " & WebHelpers.PrepareTextForShell(Me.Username) & ":" & WebHelpers.PrepareTextForShell(Me.Password)
End Sub
```

```vb
' Api.bas
Dim Client As New WebClient
Dim Authenticator As New HttpBasicAuthenticator

Authenticator.Username = "Tim"
Authenticator.Password = "password"
Set Client.Authenticator = Authenticator
```

After:

```vb
' HttpBasicAuthenticator.cls
Private WithEvents pClient As WebClient

Private Const web_HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0

Public Username As String
Public Password As String

Public Sub ConnectTo(Client As WebClient)
  Set pClient = Client
End Sub

Private Sub pClient_BeforeExecute(ByRef Request As WebRequest)
    Request.SetHeader "Authorization", "Basic " & WebHelpers.Base64Encode(Me.Username & ":" & Me.Password)
End Sub

Private Sub pClient_PrepareHttp(Request As WebRequest, ByRef Http As Object)
    Http.SetCredentials Me.Username, Me.Password, web_HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
End Sub

Private Sub pClient_PrepareCurl(Request As WebRequest, ByRef Curl As String)
    Curl = Curl & " --basic --user " & WebHelpers.PrepareTextForShell(Me.Username) & ":" & WebHelpers.PrepareTextForShell(Me.Password)
End Sub
```

```vb
' Api.bas
Dim Client As New WebClient
Dim Authenticator As New HttpBasicAuthenticator

Authenticator.Username = "Tim"
Authenticator.Password = "password"
Authenticator.ConnectTo Client
```

The following is an example of adding custom functionality (a "CookieJar") to VBA-Web without having to directly extend any of the core functionality:

```vb
' CookieJar.cls
Private WithEvents pClient As WebClient

Public Cookies As Collection

Public Sub ConnectTo(Client As WebClient)
  Set pClient = Client
End Sub

Private Sub pClient_BeforeExecute(ByRef Request As WebRequest)
  ' -> Add cookies to request
End Sub

Private Sub pClient_Response(Request As WebRequest, Response As WebResponse)
  ' <- Get cookies from response
End Sub

Private Sub Class_Initialize()
  Set Me.Cookies = New Collection
End Sub
```

```vb
' Api.bas
Dim Client As New WebClient
Dim Jar As New CookieJar

Jar.ConnectTo Client
```

# How We Teach This

This is inherently an advanced feature, but in the future some foundational features (e.g. async requests) may be built off this, so it warrants an in-depth guide for using events with VBA-Web.

Need to look into how this interacts with `Application.EnableEvents = False`. That may cause some unexpected issues when using extensions that users should be aware of.

# Drawbacks

Consumers of events have to be classes. Currently, custom formatters may be `Public` functions in `Modules`, but with an evented approach, they would need to `Classes` in order to use `WithEvents` with the `WebClient`. This is an acceptable tradeoff with the advanced functionality that this will open up to VBA-Web.

# Alternatives

An alternative would be to expand `IWebAuthenticator` to a more general `IWebExtension` interface that includes all of the potential extension points for VBA-Web. I believe this is an undue burden that would only really be needed if one of the extension points expected a return value from the extension, which is not the case in the current design. Currently, `IWebAuthenticator` uses `ByRef` to update client, requests, and responses and this could still be done with Events.

# Unresolved questions

This can be added in a backwards-compatible way for v4, with the expectation that `Client.Authenticator` and `IAuthenticator` will be deprecated. The custom formatter system includes content-type and custom functionality that occurs during parsing so removing that would lose quite a bit of functionality. It's recommended to keep that system in place, but there is room to investigate it being redesigned to be built off the Event system.
