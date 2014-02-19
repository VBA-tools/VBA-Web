Excel-REST: Excel REST Client
=============================

It seems like everything has a REST webservice these days and there is no reason to leave Excel out of the fun. Also, as of V0.2 there's async support!

Getting started
---------------

1.  In a new or existing workbook, open VBA (Alt+F11) and import all files from the src/ directory into the workbook (`RestClientBase.bas` is optional)
2.  Add a reference to Microsoft Scripting Runtime: VBA Window > Tools > References > Select Microsoft Scripting Runtime
3.  In a new module or class, create a new RestClient for the service, create a new RestRequest to request something specific from the service,
    and then use the client to execute the request
    (See below for a simple example)
4.  That's it! There are many advanced uses for Excel-REST, including asynchronous requests so that Excel isn't locked up, Authenticators for accessing
    services with Basic, OAuth1, and OAuth2 authentication, and detailed requests for complex APIs. Find out more in the [Wiki](https://github.com/timhall/Excel-REST/wiki)

The first step can be tedious, so you may want to use the blank workbook provided with the project.

Example
-------

The following is a simple of example of calling the Google Maps API to get directions between two locations (including the travel time) and processing the results

### Simple Example: Get directions
```VB
Function GetDirections(Origin As String, Destination As String) As String
    ' Create a RestClient for executing requests
    ' and set a base url that all requests will be appended to
    Dim MapsClient As New RestClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
    
    ' Create a RestRequest for getting directions
    Dim DirectionsRequest As New RestRequest
    DirectionsRequest.Resource = "directions/{format}"
    DirectionsRequest.Method = httpGET
    
    ' Set the request format -> Sets {format} segment, content-types, and parses the response
    DirectionsRequest.Format = json
    
    ' (Alternatively, replace {format} segment directly)
    DirectionsRequest.AddUrlSegment "format", "json"
    
    ' Add parameters to the request (as querystring for GET calls and body otherwise)
    DirectionsRequest.AddParameter "origin", Origin
    DirectionsRequest.AddParameter "destination", Destination
    
    ' Force parameter as querystring for all requests
    DirectionsRequest.AddQuerystringParam "sensor", "false"
    
    ' => GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false
    
    ' Execute the request and work with the response
    Dim Response As RestResponse
    Set Response = MapsClient.Execute(DirectionsRequest)
    
    If Response.StatusCode = 200 Then
        ' Work directly with parsed json data
        Dim Route As Object
        Set Route = Response.Data("routes")(1)("legs")(1)
        
        GetDirections = "It will take " & Route("duration")("text") & _
            " to travel " & Route("distance")("text") & _
            " from " & Route("start_address") & _
            " to " & Route("end_address")
    Else
        GetDirections = "Error: " & Response.Content
    End If
End Function
```

### Async Example: Add async to getting directions
```VB
Function GetDirections(Origin As String, Destination As String) As String
    Dim MapsClient As New RestClient
    Dim DirectionsRequest As New RestRequest
    ' ... Create client and request using Simple Example
    
    ' Execute the request asynchronously with RestResponse being passed into callback
    MapsClient.ExecuteAsync DirectionsRequest, "ProcessDirections"
    
    ' Keep working, handling response later
End Function

Public Sub ProcessDirections(Response As RestResponse)
    ' Handle response once the request has returned
    If Response.StatusCode = Ok Then
        Dim Route As Dictionary
        Set Route = Response.Data("routes")(1)("legs")(1)

        Debug.Print "It will take " & Route("duration")("text") & _
            " to travel " & Route("distance")("text") & _
            " from " & Route("start_address") & _
            " to " & Route("end_address")
    Else
        Debug.Print "Error: " & Response.Content
    End If
End Sub
```

For more details, check out the [Wiki](https://github.com/timhall/Excel-REST/wiki)

### Release Notes

#### 2.2.0

- Add cookies support with `Request.AddCookie(key, value)` and `Response.Cookies`
- __2.2.1__ Add `Response.Headers` collection of response headers

#### 2.1.0

- Add Microsoft Scripting Runtime dependency (for Dictionary support)
- Add `RestClient.SetProxy` for use in proxy environments
- __2.1.1__ Use `Val` for number parsing in locale-dependent settings
- __2.1.2__ Add raw binary `Body` to `RestResponse` for handling files (thanks [@berkus](https://github.com/berkus))
- __2.1.3__ Bugfixes and refactor

#### 2.0.0

- Remove JSONLib dependency (merged with RestHelpers)
- Add RestClientBase for future use with extension for single-client applications
- Add build scripts for import/export
- New specs and bugfixes
- __2.0.1__ Handle duplicate keys when parsing json
- __2.0.2__ Add Content-Length header and 408 status code for timeout

#### 1.1.0

Major Changes:

- Integrate Excel-TDD to fully test Excel-REST library
- Handle timeouts for sync and async requests
- Remove reference dependencies and use CreateObject instead

Bugfixes:

- Add cachebreaker as querystring param only
- Add Join helpers to resolve double-slash issue between base and resource url
- Only add "?" for querystring if querystring will be created and "?" isn't present
- Only put parameters in body if there are parameters

#### 0.2

- Add async support

### About

- Design based heavily on the awesome [RestSharp](http://restsharp.org/)
- Author: Tim Hall
- License: MIT
[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/304523f72ecef00eae1840dcac0c16bd "githalytics.com")](http://githalytics.com/timhall/Excel-REST)
