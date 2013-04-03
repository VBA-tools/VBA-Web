Excel-REST: Excel REST Client
=============================

It seems like everything has a REST webservice these days and there is no reason to leave Excel out of the fun. Also, as of V0.2 there's async support!

Getting started
---------------

1.  In a new or existing workbook, open VBA (Alt+F11) and import all files from the src/ directory into the project as well as JSONLib from the lib/ directory
2.  Add references (Tools > References in VBA) to Microsoft Scripting Runtime and Microsoft XML, v3.0 or above
3.  In a new module or class, create a new RestClient for the service, create a new RestRequest to request something specific from the service,
    and then use the client to execute the request
    (See below for a simple example)
4.  That's it! There are many advanced uses for Excel-REST, including asynchronous requests so that Excel isn't locked up, Authenticators for accessing
    services with Basic, OAuth1, and OAuth2 authentication, and detailed requests for complex APIs. Find out more in the [Wiki](https://github.com/timhall/Excel-REST/wiki)

Steps 1 and 2 are tedious, so you may want to use the blank workbook provided with the project. Design based heavily on the awesome [RestSharp](http://restsharp.org/)

Example
-------

The following is a simple of example of calling the Google Maps API to get directions between two locations (including the travel time) and processing the results

### Create a RestClient
```VB
Function MapsClient() As RestClient
    Set MapsClient = New RestClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
    ' (all requests will be appended to this)
    
End Function
```

### Create a RestRequest
```VB
Function DirectionsRequest(Origin As String, Destination As String) As RestRequest
    Set DirectionsRequest = New RestRequest
    DirectionsRequest.Resource = "directions/{format}"
    
    ' Set the request format (Set {format} segment, content-types, and parse the response)
    DirectionsRequest.Format = json
    
    ' Replace any {...} url segments
    ' e.g. Resource = "resource/{id}"
    ' Request.AddUrlSegment("id", 123) -> "resource/123"
    
    ' Add parameters to the request (querystring for GET calls and body for everything else)
    DirectionsRequest.AddParameter "origin", Origin
    DirectionsRequest.AddParameter "destination", Destination
    ' (or force as querystring)
    DirectionsRequest.AddQuerystringParam "sensor", "false"
    
    ' (GET, POST, PUT, DELETE, PATCH)
    DirectionsRequest.Method = httpGET
    
    ' => GET https://maps.googleapis.com/maps/api/directions/json?origin=...&destination=...&sensor=false
    
End Function
```

### Execute and process a request
```VB
Sub GetDirections()
    ' Execute the request asynchonously and process later
    MapsClient.ExecuteAsync DirectionsRequest("Raleigh, NC", "San Francisco, CA"), "ProcessDirections"
End Sub

Public Sub ProcessDirections(Response As RestResponse)
    If Response.StatusCode = Ok Then
        Dim Route As Dictionary 
        Set Route = Response.Data("routes")(1)("legs")(1)
    
        Debug.Print "It will take " & Route("duration")("text") & _
            " to travel " & Route("distance")("text") & _
            " from " & Route("start_address") & _
            " to " & Route("end_address")
    End If
End Sub
```

For more details, check out the [Wiki](https://github.com/timhall/Excel-REST/wiki)

Author: Tim Hall

License: MIT
