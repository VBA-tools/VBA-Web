Excel-REST: Excel REST Client
=============================

It seems like everything has a REST webservice these days and there is no reason to leave Excel out of the fun. Also, as of V0.2 there's async support!

Getting started
---------------

1.  In a new or existing workbook, open VBA (Alt+F11) and import all files from the src/ directory into the project as well as JSONLib from lib/ directory
    (This part is tedious, so you may want to use the blank workbook provided with the project)
2.  In a new module or class, create a new RestClient for the service, create a new RestRequest to request something specific from the service,
    and then use the client to execute the request, storing the RestResponse
    (See below for a simple example)
3.  That's it! There are many advanced uses for Excel-REST, including asynchronous requests so that Excel isn't locked up, Authenticators for accessing
    a variety of protected resources, and detailed requests for complex APIs. Find out more in the [Wiki](https://github.com/timhall/Excel-REST/wiki)

(Design based heavily on the awesome [RestSharp](http://restsharp.org/))

Examples
--------

Create a RestClient
```VB
Function MapsClient() As RestClient
    Set MapsClient = New RestClient
    
    ' Set the base url for the service
    '
    ' All requests will be appended to the base url
    ' e.g.
    ' BaseUrl = https://api.service.com/
    ' RequestA -> https://api.service.com/RequestA
    ' RequestB -> https://api.service.com/RequestB
    '
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
End Function
```

Create a RestRequest
```VB
Function DirectionsRequest(Origin As String, Destination As String) As RestRequest
    Set DirectionsRequest = New RestRequest
    
    ' Set the resource for the request
    ' 1. This will be appended to the base url of the client
    ' 2. Any segments surrounded by {...} can be replaced later
    '    {format} is a special segment that will be replaced when the format is defined
    DirectionsRequest.Resource = "{action}/{format}"
    
    ' Set the request format
    ' - This sets the {format} segment and parses the response into the desired format
    DirectionsRequest.Format = json
    
    ' Replace any url segments
    DirectionsRequest.AddUrlSegment "action", "directions"
    
    ' Add a parameter to the request
    ' (Added as a querystring for GET calls and added to the request body for everything else)
    DirectionsRequest.AddParameter "origin", Origin
    DirectionsRequest.AddParameter "destination", Destination
    
    ' Add a querystring parameter to the request
    ' (Just like AddParameter only forces it to querystring)
    DirectionsRequest.AddQuerystringParam "sensor", "false"
    
    ' Set request method
    ' (GET, POST, PUT, DELETE, PATCH)
    DirectionsRequest.Method = httpGET
    
End Function
```

Execute a request
```VB
Sub GetDirections()
    ' Execute the request synchronously and store the response
    Dim Response As RestResponse
    Set Response = MapsClient.Execute(DirectionsRequest("Cary, NC", "Raleigh, NC"))
    
    ' Execute the request asynchonously and process later
    MapsClient.ExecuteAsync DirectionsRequest("Raleigh, NC", "Cary, NC"), "ProcessDirections"
End Sub

Public Sub ProcessDirections(Response As RestResponse)
    If Response.StatusCode = Ok Then
        Dim Route As Dictionary
        Dim Duration As String
        Dim Distance As String
        Dim StartAddress As String
        Dim EndAddress As String
        
        Set Route = Response.Data("routes")(1)("legs")(1)
        Duration = Route("duration")("text")
        Distance = Route("distance")("text")
        StartAddress = Route("start_address")
        EndAddress = Route("end_address")
    
        Debug.Print "It will take " & Duration & " to travel " & Distance & " from " & StartAddress & " to "; EndAddress
    End If
End Sub
```

For more advanced examples, check out the [Wiki](https://github.com/timhall/Excel-REST/wiki)

Author: Tim Hall
License: MIT
