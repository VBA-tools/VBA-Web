---
id: get-request
title: Making a GET Request
permalink: get-request.html
prev: index.html
next: setting-up-restclient.html
---

# Making a GET Request

Excel-REST consists of three primary components: 

- `RestClient` makes requests and handles responses
- `RestRequest` define the request
- `RestResponse` format the response

## 1. Create Client

Instructions...

```VB.net{2-4}
Sub GetRequest()
    ' Create Client
    Dim MapsClient As New RestClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
End Sub
```