Option Explicit

Public Function HttpbinLookup(RequestUrl As String, Post_data As String) As WebResponse
    
    '------------------ build webclient ------------------
    Dim HttpbinClient As New WebClient
    HttpbinClient.BaseUrl = "https://httpbin.org/"
    
    
    '------------------ http basic authentication ------------------
    If Left(RequestUrl, 10) = "basic-auth" Then
        Dim HttpbinAuth1 As New HttpBasicAuthenticator       'calls setup sub inside class module HttpBasicAuthenticator
        'enter your username and password below
        HttpbinAuth1.Setup _
            Username:="user", _
            Password:="passwd"
        'add the info from the authenticator to the webclient we just created
        Set HttpbinClient.Authenticator = HttpbinAuth1
    
    
    '------------------ http digest authentication ------------------
    ElseIf Left(RequestUrl, 11) = "digest-auth" Then
        Dim HttpbinAuth2 As New DigestAuthenticator          'calls setup sub inside class module DigestAuthenticator
        'enter your username and password below
        HttpbinAuth2.Setup _
            Username:="user", _
            Password:="passwd"
        'add the info from the authenticator to the webclient we just created
        Set HttpbinClient.Authenticator = HttpbinAuth2
        'httpbin digest auth will not work without a cookie!
    End If
    
    
    '------------------ build query url request (->) ------------------
    Dim request As New WebRequest
    request.Resource = RequestUrl                            'adds request onto end of the baseurl
    'Request.AddQuerystringParam "key", Credentials.Values("Google")("api_key")                  'looks in credentials text file
    'Request.AddQuerystringParam "Request", Post_data        'outputs ?Request=Post_data          Post_data is value from cell B2
    
    
    '------------------ set formatting ------------------
    'Simple - send and receive in the same format
    'Request.Format = WebFormat.Json                         'Request.Format sets four things:   Content-Type header         Accept header
                                                                                                'Request Body conversion     Response Data conversion
    'Medium - send and receive in two different formats
    request.RequestFormat = WebFormat.JSON                   'Set Content-Type and request converter
    request.ResponseFormat = WebFormat.JSON                  'Set Accept and response converter
    'request.ResponseFormat = WebFormat.FormUrlEncoded
    
    'Advanced: Set separate everything
    'Request.RequestFormat = WebFormat.Json
    'Request.ContentType = "application/json"
    'Request.ResponseFormat = WebFormat.Json
    'Request.Accept = "application/json"
    
    
    '------------------ set method ------------------
    If RequestUrl = "post" Then
        request.Method = WebMethod.HttpPost                  'POST - form data appears within the message body of the HTTP request, not in the URL
        'Request.Body = "{""a"":123,""b"":[456, 789]}"       'same as below just all in one line
        'Request.AddBodyParameter "a", 123
        'Request.AddBodyParameter "b", Array(456, 789)
        Dim system_time As String
        system_time = Now()
        request.AddBodyParameter "systemtime", system_time   'send current system time
        request.AddBodyParameter "postdata", Post_data       'Post_data is value passed from cell B2
    Else
        request.Method = WebMethod.HttpGet                   'GET - all form data is encoded into the URL - less flexible, less secure
    End If
    
    
    '------------------ set contents ------------------
    'Add other things common to all Requests
    request.AddCookie "cookie", "testCookie"                 'httpbin digest auth will not work without a cookie!
    request.AddHeader "header", "testHeader"
    
    
    '------------------ send request and receive response ------------------
    'this takes the HttpbinClient webclient we built at the top and executes the Request webrequest we made
    'then it sets the function to return the data from the server (<-)
    Set HttpbinLookup = HttpbinClient.Execute(request)       'goes to WebClient(Execute)
   
End Function


'this is just for testing in debug window and bypasses using an excel worksheet
Public Sub Test()
    
    WebHelpers.EnableLogging = True                          'extended debug info
    
    Dim Response As WebResponse
    Set Response = HttpbinLookup("ip", "")                   'this calls function HttpbinLookup above
  
    If Response.StatusCode = WebStatusCode.OK Then
        Debug.Print "Result: " & Response.Data("origin")     'ip address
    Else
        Debug.Print Response.Content
    End If
End Sub
