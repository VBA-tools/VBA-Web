Attribute VB_Name = "Maps"
Function GetDirections(Origin As String, Destination As String) As String
    ' Create a WebClient for executing requests
    ' and set a base url that all requests will be appended to
    Dim MapsClient As New WebClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
    
    ' Create a WebRequest for getting directions
    Dim DirectionsRequest As New WebRequest
    DirectionsRequest.Resource = "directions/{format}"
    DirectionsRequest.Method = HttpGet
    
    ' Set the request format -> Sets {format} segment, content-types, and parses the response
    DirectionsRequest.Format = Json
    
    ' (Alternatively, replace {format} segment directly)
    DirectionsRequest.AddUrlSegment "format", "json"
    
    ' Add parameters to the request (as querystring for GET calls and body otherwise)
    DirectionsRequest.AddQuerystringParam "origin", Origin
    DirectionsRequest.AddQuerystringParam "destination", Destination
    
    ' Force parameter as querystring for all requests
    DirectionsRequest.AddQuerystringParam "sensor", "false"
    
    ' => GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false
    
    ' Execute the request and work with the response
    Dim Response As WebResponse
    Set Response = MapsClient.Execute(DirectionsRequest)
    
    If Response.StatusCode = WebStatusCode.Ok Then
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

Public Sub ProcessDirections(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.Ok Then
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

Function MapsClient() As WebClient
    Set MapsClient = New WebClient

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

Function DirectionsRequest(Origin As String, Destination As String) As WebRequest
    Set DirectionsRequest = New WebRequest
    DirectionsRequest.Resource = "directions/{format}"

    ' Set the request format (Set {format} segment, content-types, and parse the response)
    DirectionsRequest.Format = WebFormat.Json
    DirectionsRequest.AddUrlSegment "format", "json"

    ' Replace any {...} url segments
    ' e.g. Resource = "resource/{id}"
    ' Request.AddUrlSegment("id", 123) -> "resource/123"

    ' Add parameters to the request (querystring for GET calls and body for everything else)
    DirectionsRequest.AddQuerystringParam "origin", Origin
    DirectionsRequest.AddQuerystringParam "destination", Destination
    ' (or force as querystring)
    DirectionsRequest.AddQuerystringParam "sensor", "false"

    ' (GET, POST, PUT, DELETE, PATCH)
    DirectionsRequest.Method = WebMethod.HttpGet

    ' => GET https://maps.googleapis.com/maps/api/directions/json?origin=...&destination=...&sensor=false
End Function

