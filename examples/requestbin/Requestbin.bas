Option Explicit

Public Function RequestbinLookup(Tempurl As String, Post_data As String) As WebResponse
    
    '------------------ build webclient ------------------
    Dim RequestbinClient As New WebClient
    RequestbinClient.BaseUrl = "https://requestb.in/"
    
    
    '------------------ build query url request (->) ------------------
    Dim request As New WebRequest
    request.Resource = Tempurl                               'adds request onto end of the baseurl
    'Request.AddQuerystringParam "key", Credentials.Values("Google")("api_key")                  'looks in credentials text file
    'Request.AddQuerystringParam "Request", Post_data        'outputs ?Request=Post_data          Post_data is value from cell B1
    
        
    '------------------ set formatting ------------------
    'Simple - send and receive in the same format
    'Request.Format = WebFormat.Json                         'Request.Format sets four things:   Content-Type header         Accept header
                                                                                                'Request Body conversion     Response Data conversion
    'Medium - send and receive in two different formats
    request.RequestFormat = WebFormat.JSON                   'Set Content-Type and request converter
    'request.ResponseFormat = WebFormat.JSON                 'Set Accept and response converter
    request.ResponseFormat = WebFormat.FormUrlEncoded
    
    'Advanced: Set separate everything
    'Request.RequestFormat = WebFormat.Json
    'Request.ContentType = "application/json"
    'Request.ResponseFormat = WebFormat.Json
    'Request.Accept = "application/json"
    
    
    '------------------ set method ------------------
    request.Method = WebMethod.HttpPost                      'POST - form data appears within the message body of the HTTP request, not in the URL
    'Request.Method = WebMethod.HttpGet                      'GET - all form data is encoded into the URL - less flexible, less secure
    
    
    '------------------ set contents ------------------
    'Request.Body = "{""a"":123,""b"":[456, 789]}"           'same as below just all in one line
    'Request.AddBodyParameter "a", 123
    'Request.AddBodyParameter "b", Array(456, 789)
    Dim system_time As String
    system_time = Now()
    request.AddBodyParameter "system time", system_time      'send current system time
    request.AddBodyParameter "spreadsheet input", Post_data  'Post_data is value passed from cell B1
    
    
    ' Add other things common to all Requests
    request.AddCookie "cookie", "testCookie"
    request.AddHeader "header", "testHeader"
    
    
    '------------------ send request and receive response ------------------
    'this takes the RequestbinClient webclient we built at the top and executes the Request webrequest we made
    'then it sets the function to return the data from the server (<-)
    Set RequestbinLookup = RequestbinClient.Execute(request) 'now it goes to WebClient(Execute)
   
End Function


'this is just for testing in debug window and bypasses using an excel worksheet
Public Sub Test()

    Dim Tempurl As String
    Dim Post_data As String
    Tempurl = "1klwlzq1"                                     'enter your bin
    Post_data = "Test12345"                                  'enter your post data
    
    Dim Response As WebResponse
    Set Response = RequestbinLookup(Tempurl, Post_data)      'this calls function RequestbinLookup above
  
    If Response.StatusCode = WebStatusCode.OK Then
        Debug.Print "Result: " & Response.Content            'server response
    Else
        Debug.Print Response.Content
    End If
End Sub
