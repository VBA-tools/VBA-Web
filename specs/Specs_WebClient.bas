Attribute VB_Name = "Specs_WebClient"
''
' Specs_WebClient
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for WebClient
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebClient"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs

    Dim Client As New WebClient
    Dim Request As WebRequest
    Dim Response As WebResponse
    Dim Body As Dictionary
    Dim BodyToString As String
    Dim i As Integer
    Dim Options As Dictionary
    Dim XMLBody As Object
    Dim Curl As String
    
    Client.BaseUrl = HttpbinBaseUrl
    Client.TimeoutMs = 5000

    ' --------------------------------------------- '
    ' Properties
    ' --------------------------------------------- '
    
    ' BaseUrl
    ' Authenticator
    ' TimeoutMS
    ' ProxyServer
    ' ProxyUsername
    ' ProxyPassword
    ' ProxyBypassList
    
    ' Insecure
    ' --------------------------------------------- '
    With Specs.It("should not be Insecure by default")
        .Expect(Client.Insecure).ToEqual False
    End With
    
    With Specs.It("[Windows-only] Insecure should set options in WinHttpRequest")
#If Mac Then
        ' (Windows-only)
        .Expect(True).ToEqual True
#Else
        Dim Http As Object
        Set Request = New WebRequest
        
        ' WinHttpRequestOption_EnableCertificateRevocationCheck = 18
        ' WinHttpRequestOption_SslErrorIgnoreFlags = 4
        
        Set Http = Client.PrepareHttpRequest(Request)
        .Expect(Http.Option(18)).ToEqual True
        .Expect(Http.Option(4)).ToEqual 0
        
        Client.Insecure = True
        
        Set Http = Client.PrepareHttpRequest(Request)
        .Expect(Http.Option(18)).ToEqual False
        .Expect(Http.Option(4)).ToEqual 13056
#End If

        Client.Insecure = False
    End With
    
    With Specs.It("[Mac-only] Insecure should set --insecure flag in cURL")
#If Mac Then
        Set Request = New WebRequest
        
        Curl = Client.PrepareCurlRequest(Request)
        .Expect(Curl).ToNotMatch "--insecure"
        
        Client.Insecure = True
        
        Curl = Client.PrepareCurlRequest(Request)
        .Expect(Curl).ToMatch "--insecure"
#Else
        ' (Mac-only)
        .Expect(True).ToEqual True
#End If

        Client.Insecure = False
    End With
    
    ' FollowRedirects
    ' --------------------------------------------- '
    With Specs.It(" should FollowRedirects")
        Set Request = New WebRequest
        Request.Resource = "redirect/5"
        Request.Format = WebFormat.PlainText
        
        Client.FollowRedirects = True
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        
        Client.FollowRedirects = False
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 302
    End With
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ' Execute
    ' --------------------------------------------- '
    With Specs.It("Execute should set method, url, headers, cookies, and body")
        Set Request = New WebRequest
        Request.Resource = "put"
        Request.Method = WebMethod.HttpPut
        Request.AddQuerystringParam "number", 123
        Request.AddQuerystringParam "string", "abc"
        Request.AddQuerystringParam "boolean", True
        Request.AddHeader "X-Custom", "Howdy!"
        Request.AddCookie "abc", 123
        Request.RequestFormat = WebFormat.FormUrlEncoded
        Request.ResponseFormat = WebFormat.Json
        Request.AddBodyParameter "message", "Howdy!"
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        .Expect(Response.Data("url")).ToEqual "http://httpbin.org/put?number=123&string=abc&boolean=true"
        .Expect(Response.Data("headers")("X-Custom")).ToEqual "Howdy!"
        .Expect(Response.Data("headers")("Content-Type")).ToMatch WebHelpers.FormatToMediaType(WebFormat.FormUrlEncoded)
        .Expect(Response.Data("headers")("Accept")).ToMatch WebHelpers.FormatToMediaType(WebFormat.Json)
        .Expect(Response.Data("headers")("Cookie")).ToMatch "abc=123"
        .Expect(Response.Data("form")("message")).ToEqual "Howdy!"
    End With

    With Specs.It("Execute should work with each method")
        Set Request = New WebRequest
        
        Request.Method = WebMethod.HttpGet
        Request.Resource = "get"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        
        Request.Method = WebMethod.HttpPost
        Request.Resource = "post"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        
        Request.Method = WebMethod.HttpPatch
        Request.Resource = "patch"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        
        Request.Method = WebMethod.HttpPut
        Request.Resource = "put"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        
        Request.Method = WebMethod.HttpDelete
        Request.Resource = "delete"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        
        Request.Method = WebMethod.HttpHead
        Request.Resource = "get"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
    End With
    
    ' GetJson
    ' --------------------------------------------- '
    With Specs.It("should GetJSON")
        Set Response = Client.GetJson("/get")

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("headers").Count).ToBeGT 0
    End With
    
    With Specs.It("should GetJSON with options")
        Set Options = New Dictionary
        Options.Add "Headers", New Collection
        Options("Headers").Add WebHelpers.CreateKeyValue("X-Custom", "Howdy!")
        Options.Add "Cookies", New Collection
        Options("Cookies").Add WebHelpers.CreateKeyValue("abc", 123)
        Options.Add "QuerystringParams", New Collection
        Options("QuerystringParams").Add WebHelpers.CreateKeyValue("message", "Howdy!")
        Options.Add "UrlSegments", New Dictionary
        Options("UrlSegments").Add "resource", "get"
        
        Set Response = Client.GetJson("/{resource}", Options)
    
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("url")).ToEqual "http://httpbin.org/get?message=Howdy!"
        .Expect(Response.Data("headers")("X-Custom")).ToEqual "Howdy!"
        .Expect(Response.Data("headers")("Cookie")).ToMatch "abc=123"
    End With
    
    ' PostJson
    ' --------------------------------------------- '
    With Specs.It("should PostJSON")
        Set Body = New Dictionary
        Body.Add "a", 3.14
        Body.Add "b", "Howdy!"
        Body.Add "c", True
        Set Response = Client.PostJson("/post", Body)

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")("a")).ToEqual 3.14
        .Expect(Response.Data("json")("b")).ToEqual "Howdy!"
        .Expect(Response.Data("json")("c")).ToEqual True

        Set Response = Client.PostJson("/post", Array(3, 2, 1))

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")(1)).ToEqual 3
        .Expect(Response.Data("json")(2)).ToEqual 2
        .Expect(Response.Data("json")(3)).ToEqual 1
    End With
    
    With Specs.It("should PostJSON with options")
        Set Body = New Dictionary
        Body.Add "a", 3.14
        Body.Add "b", "Howdy!"
        Body.Add "c", True
        
        Set Options = New Dictionary
        Options.Add "Headers", New Collection
        Options("Headers").Add WebHelpers.CreateKeyValue("X-Custom", "Howdy!")
        Options.Add "Cookies", New Collection
        Options("Cookies").Add WebHelpers.CreateKeyValue("abc", 123)
        Options.Add "QuerystringParams", New Collection
        Options("QuerystringParams").Add WebHelpers.CreateKeyValue("message", "Howdy!")
        Options.Add "UrlSegments", New Dictionary
        Options("UrlSegments").Add "resource", "post"
        
        Set Response = Client.PostJson("/{resource}", Body, Options)
    
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("url")).ToEqual "http://httpbin.org/post?message=Howdy!"
        .Expect(Response.Data("headers")("X-Custom")).ToEqual "Howdy!"
        .Expect(Response.Data("headers")("Cookie")).ToMatch "abc=123"
        .Expect(Response.Data("json")("a")).ToEqual 3.14
        .Expect(Response.Data("json")("b")).ToEqual "Howdy!"
        .Expect(Response.Data("json")("c")).ToEqual True
    End With
    
    ' SetProxy
    
    ' GetFullUrl
    ' --------------------------------------------- '
    With Specs.It("should GetFullUrl of Request")
        Set Request = New WebRequest
        
        Client.BaseUrl = "https://facebook.com/api"
        Request.Resource = "status"
        .Expect(Client.GetFullUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api"
        Request.Resource = "/status"
        .Expect(Client.GetFullUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api/"
        Request.Resource = "status"
        .Expect(Client.GetFullUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api/"
        Request.Resource = "/status"
        .Expect(Client.GetFullUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = HttpbinBaseUrl
    End With
    
    ' PrepareHttpRequest
    
    ' PrepareCURL
    ' @internal
    ' --------------------------------------------- '
    With Specs.It("[Mac-only] should PrepareCURLRequest")
#If Mac Then
        Set Client = New WebClient
        Client.BaseUrl = "http://localhost:3000/"
        Client.ProxyServer = "proxyserver"
        Client.ProxyBypassList = "proxyserver:80, *.github.com"
        Client.ProxyUsername = "proxyuser"
        Client.ProxyPassword = "proxypassword"
        
        Set Request = New WebRequest
        Request.Resource = "text"
        Request.AddQuerystringParam "type", "message"
        Request.Method = HttpPost
        Request.RequestFormat = WebFormat.PlainText
        Request.ResponseFormat = WebFormat.Json
        Request.Body = "Howdy!"
        Request.AddHeader "custom", "Howdy!"
        Request.AddCookie "test-cookie", "howdy"
        
        Curl = Client.PrepareCurlRequest(Request)
        .Expect(Curl).ToMatch "http://localhost:3000/text?type=message"
        .Expect(Curl).ToMatch "-X POST"
        .Expect(Curl).ToMatch "--proxy proxyserver"
        .Expect(Curl).ToMatch "--noproxy proxyserver:80, *.github.com"
        .Expect(Curl).ToMatch "--proxy-user proxyuser:proxypassword"
        .Expect(Curl).ToMatch "-H 'Content-Type: text/plain'"
        .Expect(Curl).ToMatch "-H 'Accept: application/json'"
        .Expect(Curl).ToMatch "-H 'custom: Howdy!'"
        .Expect(Curl).ToMatch "--cookie 'test-cookie=howdy;'"
        .Expect(Curl).ToMatch "-d 'Howdy!'"
#Else
        ' (Mac-only)
        .Expect(True).ToEqual True
#End If
    End With
    
    With Specs.It("should handle timeout errors")
        Client.TimeoutMs = 500
        
        Set Request = New WebRequest
        Request.Resource = "delay/{seconds}"
        Request.AddUrlSegment "seconds", "5"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 408
        .Expect(Response.StatusDescription).ToMatch "Request Timeout"
    End With
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    On Error Resume Next
    
    With Specs.It("should throw 11011 on general error")
        ' Unsupported protocol
        Client.BaseUrl = "unknown://"
        Set Response = Client.Execute(Request)
        
        .Expect(Err.Number).ToEqual 11011 + vbObjectError
        Err.Clear
    End With
    
    Set Client = Nothing
End Function

Public Function OfflineSpecs() As SpecSuite
    ' Disconnect from the internet before running these specs
    
    Set OfflineSpecs = New SpecSuite
    OfflineSpecs.Description = "WebClient - Offline"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs
    
    Dim Client As New WebClient
    Dim Request As WebRequest
    Dim Response As WebResponse
    
    Client.BaseUrl = HttpbinBaseUrl
    
    With OfflineSpecs.It("should handle resolve errors as timeout")
        Client.TimeoutMs = 500
        
        Set Request = New WebRequest
        Request.Resource = "/get"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 408
        .Expect(Response.StatusDescription).ToEqual "Request Timeout"
    End With
    
    With OfflineSpecs.It("should not crash with auto-proxy resolve error")
        Client.EnableAutoProxy = True
        Client.TimeoutMs = 500
        
        Set Request = New WebRequest
        Request.Resource = "/get"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 408
        .Expect(Response.StatusDescription).ToEqual "Request Timeout"
    End With
End Function

