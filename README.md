Excel Helpers
=============
A helping hand to those who like to push the limits of Excel and VBA

1. Excel REST Client
----------------------
It seems like everything has a REST webservice these days and there is no reason to leave Excel out of the fun. Also, as of V0.2 there's async support! (Example workbook coming in a jiffy)

(API design based heavily on the awesome [RestSharp](http://restsharp.org/))

Example:

```VB
Sub TwitterSearch()
	Dim twitterClient As New RestClient
	twitterClient.baseUrl = "https://api.twitter.com/{ApiVersion}/"
	
	' Setup OAuth1 authentication (there are also authenticators for HTTP Basic and OAuth2)
	Set twitterClient.Authenticator = RestModule.OAuth1( _
	    ConsumerKey:="your-consumer-key", _
	    ConsumerSecret:="your-consumer-secret", _
	    Token:="your-token", _
	    TokenSecret:="your-token-secret" _
	)
	
	Dim search As New RestRequest
	search.Resource = "search.{format}"
	
	' Set the request format
	' Replaces the format keyword in the Resource, sets the content/type, and parses the results
	search.Format = json
	
	' Add a parameter to the request
	' (Added as a querystring for GET calls and added to the request body for everything else)
	search.AddParameter "q", "Excel"
	
	' Add a url segment
	' Replaces any {tags} in the BaseUrl or Resource
	search.AddUrlSegment "ApiVersion", "1"
	
	' Execute the request (synchronous and asynchronous(!) methods available)
	Call twitterClient.ExecuteAsync(search, "HandleSearchResults")
End Sub

' Callback for async call that takes in the response from the request
Sub HandleSearchResults(response As RestResponse)
	Debug.Print _
		"This was handled asynchronously: " & _
		response.StatusCode & " (" & response.StatusDescription & "): " _
		& response.Content
End Sub
```

2. Excel Testing Library
--------------------------
Bring the reliability of other programming realms to Excel.

(API design based heavily on [Jasmine](http://pivotal.github.com/jasmine/))

```VB
'... (coming soon :))
```