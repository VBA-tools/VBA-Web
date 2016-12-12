- Start Date: 2016-12-12
- RFC PR: (leave this empty)
- VBA-Web Issue: (leave this empty)

# Summary

Currently, VBA-Web has a two mechanisms for extension: Authenticators and Custom Formatters. Authenticators have access to the http and curl objects when they're being prepared, the request before it's executed, and the response. Custom Formatters have access to the response body. These are very valuable extension points for adding more advanced functionality to VBA-Web and a more generalized system based on events would simplify these two existing extension types and allow for other customized functionality in the future.

# Motivation

The event system could provide the foundation for the two existing extension types and be used for more advanced functionality, all while simplifying the core code. `IWebAuthenticator` could be removed with Authenticators using events, more advanced Custom Formatters could be added for things other than just converting the request/response body, and it would be straightforward to add async functionality in the future.

# Detailed design

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
| -> a) PrepareHttp
| -> b) BeforeRequest
| -> c) PrepareCurl
Execute request
| -> d) Execute
Receive response
| -> e) Parse
| -> f) Response
Return response
```

### a) `PrepareHttp(ByRef Request As WebRequest, ByRef Http As WinHttpRequest)` (Windows)

Called after the `WinHttpRequest` has been prepared, this event can be used to interact with the `WinHttpRequest` directly (e.g. setting options, authentication, etc.).

### b) `BeforeRequest(ByRef Request As WebRequest)`

This event is called before the request is generated from the `WebReqest`, so that changes can be made to `WebRequest` before it is parsed and executed (e.g. setting authentication headers, custom parsing, etc.).

### c) `PrepareCurl(ByRef Request As WebReqest, ByRef Curl As String)` (Mac)

Called after the curl command has been prepared (occurs after `BeforeRequest` since the request is parsed for the curl command), this event can be used to interact with the curl command directly (e.g. adding flags).

### d) `Execute(Request As WebRequest)`

This event occurs when the request is sent, before a response has been received.

### e) `Parse(ByRef Request As WebRequest, ByRef Response As WebResponse)`

This event occurs after the response has been initially parsed, but before it has been returned, and can be used to perform any custom parsing / response handling.

### f) `Response(Request As WebRequest, Response As WebResponse)`

This event occurs after the entire request lifecycle is complete and the response parsing is finalized.

# How We Teach This

This is inherently an advanced feature, but in the future some foundational features (e.g. async requests) may be built off this, so it warrants an in-depth guide for using events with VBA-Web.

Need to look into how this interacts with `Application.EnableEvents = False`. That may cause some unexpected issues when using extensions that users should be aware of.

# Drawbacks

Consumers of events have to be classes. Currently, custom formatters may be `Public` functions in `Modules`, but with an evented approach, they would need to `Classes` in order to use `WithEvents` with the `WebClient`. This is an acceptable tradeoff with the advanced functionality that this will open up to VBA-Web.

# Alternatives

An alternative would be to expand `IWebAuthenticator` to a more general `IWebExtension` interface that includes all of the potential extension points for VBA-Web. I believe this is an undue burden that would only really be needed if one of the extension points expected a return value from the extension, which is not the case in the current design. Currently, `IWebAuthenticator` uses `ByRef` to update client, requests, and responses and this could still be done with Events.

# Unresolved questions

None.
