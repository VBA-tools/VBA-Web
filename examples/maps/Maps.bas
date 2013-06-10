Attribute VB_Name = "Maps"
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

