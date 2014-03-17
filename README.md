Excel-REST: Excel REST Client
=============================

It seems like everything has a REST webservice these days and there is no reason to leave Excel out of the fun. Excel-REST is designed to make working with complex webservices easy with Excel. It includes support for authentication, making async requests, automatically converting and parsing JSON, working with cookies and headers in requests and responses, and much more.

Getting started
---------------

1. Download the [latest release (v3.0.0)](https://github.com/timhall/Excel-REST/releases)
2. `Excel-REST - Blank.xlsm` has everything setup and ready to go.

For more details see the [Wiki](https://github.com/timhall/Excel-REST/wiki)

Examples
-------

The following examples demonstrate using the Google Maps API to get directions between two locations.

### GetJSON Example
```VB
Function GetDirections(Origin As String, Destination As String) As String
    ' Create a RestClient for executing requests
    ' and set a base url that all requests will be appended to
    Dim MapsClient As New RestClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
    
    ' Use GetJSON helper to execute simple request and work with response
    Dim Resource As String
    Dim Response As RestResponse
    
    Resource = "directions/json?origin=" & Origin & "&destination=" & Destination & "&sensor=false"
    Set Response = MapsClient.GetJSON(Resource)
    
    ' => GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false
    
    ProcessDirections Response
End Function

Public Sub ProcessDirections(Response As RestResponse)
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

There are 3 primary components in Excel-REST: 

1. `RestRequest` for defining complex requests
2. `RestClient` for executing requests
3. `RestResponse` for dealing with responses. 
 
In the above example, the request is fairly simple, so we can skip creating a `RestRequest` and instead use the `Client.GetJSON` helper to GET json from a specific url. In processing the response, we can look at the `StatusCode` to make sure the request succeeded and then use the parsed json in the `Data` parameter to extract complex information from the response. 

### RestRequest Example

If we wish to have more control over the request, the following example uses `RestRequest` to define a complex request.

```VB
Function GetDirections(Origin As String, Destination As String) As String
    Dim MapsClient As New RestClient
    ' ... Setup client using GetJSON Example
    
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
    
    ProcessDirections Response
End Function

Public Sub ProcessDirections(Response As RestResponse)
    ' ... Same as previous examples
End Sub
```

The above example demonstrates some of the powerful feature available with `RestRequest`. Some of the features include:

- Url segments (Replace {segment} in resource with value)
- Method (GET, POST, PUT, PATCH, DELETE)
- Format (json and url-encoded) for content-type and converting/parsing request and response
- Parameters and QuerystringParams
- Body
- Cookies
- Headers

For more details, see the `RestRequest` page in with [Wiki](https://github.com/timhall/Excel-REST/wiki/RestRequest)

### Async Example

The above examples execute synchronously, but Excel-REST can run them asynchronously with ease so that your program can keep working and handle the response later once the request completes.

```VB
Function GetDirections(Origin As String, Destination As String) As String
    Dim MapsClient As New RestClient
    Dim DirectionsRequest As New RestRequest
    ' ... Create client and request using RestRequest Example
    
    ' Execute the request asynchronously
    ' Pass in name of Public Sub as callback to be called asynchronously once request completes
    MapsClient.ExecuteAsync DirectionsRequest, "ProcessDirections"
    
    ' Keep working, handling response later
End Function

Public Sub ProcessDirections(Response As RestResponse)
    ' (Called asynchronously!)
    ' ... Same as previous examples
End Sub
```

### Authentication Example

The following example demonstrates using an authenticator with Excel-REST to query Twitter. The `TwitterAuthenticator` (found in the `authenticators/` [folder](https://github.com/timhall/Excel-REST/tree/master/authenticators)) uses Twitter's OAuth 1.0a authentication and details of how it was created can be found in the [Wiki](https://github.com/timhall/Excel-REST/wiki/Implementing-your-own-IAuthenticator).

```VB
Function QueryTwitter(query As String) As RestResponse
    Dim TwitterClient As New RestClient
    TwitterClient.BaseUrl = "https://api.twitter.com/1.1/"
    
    ' Setup authenticator
    Dim TwitterAuth As New TwitterAuthenticator
    TwitterAuth.Setup _
        ConsumerKey:="Your consumer key", _
        ConsumerSecret:="Your consumer secret"
    Set TwitterClient.Authenticator = TwitterAUth
    
    ' Setup query request
    Dim Request As New RestRequest
    Request.Resource = "search/tweets.{format}"
    Request.Format = json
    Request.Method = httpGET
    Request.AddParameter "q", query
    Request.AddParameter "lang", "en"
    Request.AddParameter "count", 20
    
    ' => GET https://api.twitter.com/1.1/search/tweets.json?q=...&lang=en&count=20
    '    Authorization Bearer Token... (received and added automatically via TwitterAuthenticator)
    
    Set QueryTwitter = TwitterClient.Execute(Request)
End Function
```

For more details, check out the [Wiki](https://github.com/timhall/Excel-REST/wiki) and [Examples](https://github.com/timhall/Excel-REST/tree/master/examples)

### Release Notes

#### 3.0.0

- Add `Client.GetJSON` and `Client.PostJSON` helpers to GET and POST JSON without setting up request
- Add `AfterExecute` to `IAuthenticator` (Breaking change, all IAuthenticators must implement this new method)

#### 2.3.0

- Add `form-urlencoded` format and helpers
- Combine Body + Parameters and Querystring + Parameters with priority given to Body or Querystring, respectively

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
