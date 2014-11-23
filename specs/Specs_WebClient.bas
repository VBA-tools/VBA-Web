Attribute VB_Name = "Specs_WebClient"
''
' Specs_WebClient
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Specs for WebClient
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebClient"
    
    Dim Client As New WebClient
    Dim Request As WebRequest
    Dim Response As WebResponse
    Dim Body As Dictionary
    Dim BodyToString As String
    Dim i As Integer
    Dim Options As Dictionary
    Dim XMLBody As Object
    
    On Error Resume Next
    Client.BaseUrl = HttpbinBaseUrl
    Client.TimeoutMS = 5000
    
    ' --------------------------------------------- '
    ' Properties
    ' --------------------------------------------- '
    
    ' BaseUrl
    ' Username
    ' Password
    ' Authenticator
    ' TimeoutMS
    ' ProxyServer
    ' ProxyUsername
    ' ProxyPassword
    ' ProxyBypassList
    
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
    
    With Specs.It("Execute should use Basic Authentication")
        Set Request = New WebRequest
        Request.Resource = "basic-auth/{user}/{password}"
        Request.AddUrlSegment "user", "Tim"
        Request.AddUrlSegment "password", "Secret123"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Unauthorized
        
        Client.Username = "Tim"
        Client.Password = "Secret123"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("authenticated")).ToEqual True
    End With
    
    ' "Execute should handle timeout errors"
    ' -> Handled last due to side effects from timeout

#If Mac Then
    With Specs.It("Execute should handle cURL errors")
        ' -> Similar errors are thrown with WinHttp, match those error numbers
    End With
#End If
    
    ' GetJSON
    ' --------------------------------------------- '
    With Specs.It("should GetJSON")
        Set Response = Client.GetJSON("/get")

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
        
        Set Response = Client.GetJSON("/{resource}", Options)
    
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("url")).ToEqual "http://httpbin.org/get?message=Howdy!"
        .Expect(Response.Data("headers")("X-Custom")).ToEqual "Howdy!"
        .Expect(Response.Data("headers")("Cookie")).ToMatch "abc=123"
    End With
    
    ' PostJSON
    ' --------------------------------------------- '
    With Specs.It("should PostJSON")
        Set Body = New Dictionary
        Body.Add "a", 3.14
        Body.Add "b", "Howdy!"
        Body.Add "c", True
        Set Response = Client.PostJSON("/post", Body)

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")("a")).ToEqual 3.14
        .Expect(Response.Data("json")("b")).ToEqual "Howdy!"
        .Expect(Response.Data("json")("c")).ToEqual True

        Set Response = Client.PostJSON("/post", Array(3, 2, 1))

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
        
        Set Response = Client.PostJSON("/{resource}", Body, Options)
    
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
    With Specs.It("should GetFullUrl of path")
        Client.BaseUrl = "https://facebook.com/api"
        .Expect(Client.GetFullUrl("status")).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api"
        .Expect(Client.GetFullUrl("/status")).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api/"
        .Expect(Client.GetFullUrl("status")).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api/"
        .Expect(Client.GetFullUrl("/status")).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = HttpbinBaseUrl
    End With
    
    ' GetFullRequestUrl
    ' --------------------------------------------- '
    With Specs.It("should GetFullRequestUrl of Request")
        Set Request = New WebRequest
        
        Client.BaseUrl = "https://facebook.com/api"
        Request.Resource = "status"
        .Expect(Client.GetFullRequestUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api"
        Request.Resource = "/status"
        .Expect(Client.GetFullRequestUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api/"
        Request.Resource = "status"
        .Expect(Client.GetFullRequestUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = "https://facebook.com/api/"
        Request.Resource = "/status"
        .Expect(Client.GetFullRequestUrl(Request)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = HttpbinBaseUrl
    End With
    
    ' PrepareHttpRequest
    
    ' PrepareCURL
    ' @internal
    ' --------------------------------------------- '
#If Mac Then
    With Specs.It("should PrepareCURLRequest")
        Set Client = New WebClient
        Client.BaseUrl = "http://localhost:3000/"
        Client.Username = "user"
        Client.Password = "password"
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
        
        Dim cURL As String
        
        cURL = Client.PrepareCURLRequest(Request)
        .Expect(cURL).ToMatch "http://localhost:3000/text?type=message"
        .Expect(cURL).ToMatch "-X POST"
        .Expect(cURL).ToMatch "--user user:password"
        .Expect(cURL).ToMatch "--proxy proxyserver"
        .Expect(cURL).ToMatch "--noproxy proxyserver:80, *.github.com"
        .Expect(cURL).ToMatch "--proxy-user proxyuser:proxypassword"
        .Expect(cURL).ToMatch "-H 'Content-Type: text/plain'"
        .Expect(cURL).ToMatch "-H 'Accept: application/json'"
        .Expect(cURL).ToMatch "-H 'custom: Howdy!'"
        .Expect(cURL).ToMatch "--cookie 'test-cookie=howdy;'"
        .Expect(cURL).ToMatch "-d 'Howdy!'"
    End With
#End If
    
    With Specs.It("Execute should handle timeout errors")
        Client.TimeoutMS = 500
        
        Set Request = New WebRequest
        Request.Resource = "delay/{seconds}"
        Request.AddUrlSegment "seconds", "2"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 408
        .Expect(Response.StatusDescription).ToEqual "Request Timeout"
    End With
    
    ' Move to WebResponse
'    With Specs.It("should return status code and status description from request")
'        Set Request = New WebRequest
'        Request.Resource = "status/{code}"
'        Request.ResponseFormat = WebFormat.plaintext
'
'        Request.AddUrlSegment "code", 200
'        Set Response = Client.Execute(Request)
'        .Expect(Response.StatusCode).ToEqual 200
'        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "OK"
'
'        Request.AddUrlSegment "code", 304
'        Set Response = Client.Execute(Request)
'        .Expect(Response.StatusCode).ToEqual 304
'        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NOT MODIFIED"
'
'        Request.AddUrlSegment "code", 404
'        Set Response = Client.Execute(Request)
'        .Expect(Response.StatusCode).ToEqual 404
'        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NOT FOUND"
'
'        Request.AddUrlSegment "code", 500
'        Set Response = Client.Execute(Request)
'        .Expect(Response.StatusCode).ToEqual 500
'        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "INTERNAL SERVER ERROR"
'    End With
'
'    With Specs.It("should parse request response")
'        Set Request = New WebRequest
'        Request.Resource = "post"
'        Request.Method = httpPOST
'        Request.AddBodyParameter "a", "1"
'        Request.AddBodyParameter "b", 2
'        Request.AddBodyParameter "c", 3.14
'        Request.AddBodyParameter "d", False
'        Request.AddBodyParameter "e", Array(1)
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Data).ToNotBeUndefined
'        .Expect(Response.Data("json")("a")).ToEqual "1"
'        .Expect(Response.Data("json")("b")).ToEqual 2
'        .Expect(Response.Data("json")("c")).ToEqual 3.14
'        .Expect(Response.Data("json")("d")).ToEqual False
'        .Expect(Response.Data("json")("e")(1)).ToEqual 1
'    End With
'
'    With Specs.It("should include binary body in response")
'        Set Request = New WebRequest
'        Request.Resource = "robots.txt"
'        Request.ResponseFormat = WebFormat.plaintext
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Body).ToNotBeUndefined
'
'        If Not IsEmpty(Response.Body) Then
'            For i = LBound(Response.Body) To UBound(Response.Body)
'                BodyToString = BodyToString & Chr(Response.Body(i))
'            Next i
'        End If
'
'        .Expect(BodyToString).ToEqual "User-agent: *" & vbLf & "Disallow: /deny" & vbLf
'    End With
'
'    With Specs.It("should include headers in response")
'        Set Request = New WebRequest
'        Request.Resource = "response-headers"
'        Request.AddQuerystringParam "X-Custom", "Howdy!"
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Headers.Count).ToBeGTE 1
'
'        Dim Header As Dictionary
'        Dim CustomValue As String
'        For Each Header In Response.Headers
'            If Header("Key") = "X-Custom" Then
'                CustomValue = Header("Value")
'            End If
'        Next Header
'
'        .Expect(CustomValue).ToEqual "Howdy!"
'    End With
'
'    With Specs.It("should include cookies in response")
'        Set Request = New WebRequest
'        Request.Resource = "response-headers"
'        Request.AddQuerystringParam "Set-Cookie", "a=abc"
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Cookies.Count).ToEqual 1
'        .Expect(WebHelpers.FindInKeyValues(Response.Cookies, "a")).ToEqual "abc"
'    End With
'
'#If Mac Then
'#Else
'    With Specs.It("should convert and parse XML")
'        Set Request = New WebRequest
'        Request.Resource = "Xml"
'        Request.Format = WebFormat.Xml
'        Request.Method = httpGET
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Content).ToMatch "<slideshow"
'        .Expect(Response.Data).ToNotBeUndefined
'        .Expect(Response.Data.ChildNodes(2).ChildNodes(1).ChildNodes(0).Text).ToEqual "Wake up to WonderWidgets!"
'
'        Set XMLBody = CreateObject("MSXML2.DOMDocument")
'        XMLBody.Async = False
'        XMLBody.LoadXML "<Point><X>1.23</X><Y>4.56</Y></Point>"
'        Set Request.Body = XMLBody
'        .Expect(Request.Body).ToEqual "<Point><X>1.23</X><Y>4.56</Y></Point>"
'    End With
'#End If
    
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function

