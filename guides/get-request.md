---
id: get-request
title: Making a GET Request
permalink: get-request.html
prev: index.html
next: setting-up-restclient.html
---
<section class="docs-single">

# Making a GET Request

Excel-REST consists of three primary components: 

- `RestClient` makes requests and handles responses
- `RestRequest` define the request
- `RestResponse` format the response

</section>

<section class="docs-split">
  <div class="instructions">

## 1. Create Client

Instructions...
  </div>
  <div class="code">

```VB.net
Sub GetRequest()
    ' Create Client
    Dim MapsClient As New RestClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
End Sub
```
  </div>
</section>

<section class="docs-split">
  <div class="instructions">

## 2. Create Request

Instructions...
  </div>
  <div class="code">

```VB.net
Sub GetRequest()
    ' Create Client
    ' ...

    ' Create Request
    Dim DirectionsRequest As New RestRequest
    DirectionsRequest.Resource = "directions/{format}"
    DirectionsRequest.Method = httpGE

    ' Set the request forma
    DirectionsRequest.Format = AvailableTypes.json

    ' (Alternatively, replace {format} segment directly)
    DirectionsRequest.AddUrlSegment "format", "json"

    ' Add parameters to the request 
    ' -> querystring for GET calls and body otherwise
    DirectionsRequest.AddParameter "origin", Origin
    DirectionsRequest.AddParameter "destination", Destination

    ' Force parameter as querystring for all requests
    DirectionsRequest.AddQuerystringParam "sensor", "false"
End Sub
```
  </div>
</section>

<section class="docs-split">
  <div class="instructions">

## 3. Execute Request

Instructions...

`GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false`
  </div>
  <div class="code">

```VB.net
Sub GetRequest()
    ' Create Client
    ' ...

    ' Create Request
    ' ...

    ' Execute Request
    Dim Response As RestResponse
    Set Response = MapsClient.Execute(DirectionsRequest)
End Sub
```
  </div>
</section>
<section class="docs-split">
  <div class="instructions">

## 4. Handle Response

```json
{
  "routes": [
    {
      "legs": [
        {
          "duration": {
            "text": "..."
          },
          "distance": {
            "text": "..."
          },
          "start_address": "...",
          "end_address": "..."
        }
      ]
    }
  ]
}
```
  </div>
  <div class="code">

```VB.net
Sub GetRequest()
    ' Create Client
    ' ...

    ' Create Request
    ' ...

    ' Execute Request
    ' ...

    ' Handle Response
    Response Response.StatusCode = Ok Then
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
  </div>
</section>