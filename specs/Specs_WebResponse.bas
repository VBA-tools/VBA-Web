Attribute VB_Name = "Specs_WebResponse"
''
' Specs_WebResponse
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for WebResponse
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebResponse"
    
    Dim Client As New WebClient
    Dim Request As WebRequest
    Dim Response As WebResponse
    Dim UpdatedResponse As WebResponse
    Dim ResponseHeaders As String
    Dim Headers As Collection
    Dim Cookies As Collection
    Dim Curl As String
    
    Dim web_CrLf As String
    web_CrLf = VBA.Chr$(13) & VBA.Chr$(10)
    
    Client.BaseUrl = HttpbinBaseUrl
    Client.TimeoutMs = 5000
    
    ' --------------------------------------------- '
    ' Properties
    ' --------------------------------------------- '
    
    ' StatusCode
    ' --------------------------------------------- '
    With Specs.It("should extract status code from response")
        Set Request = New WebRequest
        Request.Resource = "status/{code}"
        
        Request.AddUrlSegment "code", 204
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.NoContent
        
        Request.AddUrlSegment "code", 304
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.NotModified
        
        Request.AddUrlSegment "code", 404
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.NotFound
    End With
    
    ' StatusDescription
    ' --------------------------------------------- '
    With Specs.It("should extract status description from response")
        Set Request = New WebRequest
        Request.Resource = "status/{code}"
        
        Request.AddUrlSegment "code", 204
        Set Response = Client.Execute(Request)
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NO CONTENT"
        
        Request.AddUrlSegment "code", 304
        Set Response = Client.Execute(Request)
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NOT MODIFIED"
        
        Request.AddUrlSegment "code", 404
        Set Response = Client.Execute(Request)
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NOT FOUND"
    End With
    
    ' Content
    ' --------------------------------------------- '
    With Specs.It("should extract unconverted content string from response")
        Set Request = New WebRequest
        Request.Resource = "user-agent"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Content).ToMatch """user-agent"":"
    End With
    
    ' Data
    ' --------------------------------------------- '
    With Specs.It("Data should use ResponseFormat to convert Content")
        Set Request = New WebRequest
        Request.Resource = "get?message=Howdy!"
        Request.ResponseFormat = WebFormat.Json
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data.Exists("headers")).ToEqual True
        .Expect(Response.Data("args")("message")).ToEqual "Howdy!"
    End With
    
    With Specs.It("Data should be nothing for PlainText")
        Set Request = New WebRequest
        Request.Resource = "get?message=Howdy!"
        Request.ResponseFormat = WebFormat.PlainText
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.Data).ToBeNothing
    End With
    
    ' Body
    ' --------------------------------------------- '
    With Specs.It("should extract raw binary bytes from response")
        Set Request = New WebRequest
        Request.Resource = "bytes/10"
        Request.ResponseFormat = WebFormat.PlainText
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.Body).ToNotBeUndefined
        .Expect(UBound(Response.Body)).ToEqual 9
    End With
    
    ' Headers
    ' --------------------------------------------- '
    With Specs.It("should extract headers from response")
        Set Request = New WebRequest
        Request.Resource = "response-headers"
        Request.AddQuerystringParam "X-Custom", "Howdy!"
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.Headers.Count).ToBeGTE 1
        .Expect(WebHelpers.FindInKeyValues(Response.Headers, "X-Custom")).ToEqual "Howdy!"
    End With
    
    ' Cookies
    ' --------------------------------------------- '
    With Specs.It("should extract cookies from response")
        Set Request = New WebRequest
        Request.Resource = "response-headers"
        Request.AddQuerystringParam "Set-Cookie", "message=Howdy!"
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.Cookies.Count).ToBeGT 0
        .Expect(WebHelpers.FindInKeyValues(Response.Cookies, "message")).ToEqual "Howdy!"
    End With
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ' Update
    ' --------------------------------------------- '
    With Specs.It("should update response")
        Set Response = New WebResponse
        Set UpdatedResponse = New WebResponse

        Response.StatusCode = 401
        Response.Body = Array("Unauthorized")
        Response.Content = "Unauthorized"

        UpdatedResponse.StatusCode = 200
        UpdatedResponse.Body = Array("Ok")
        UpdatedResponse.Content = "Ok"

        Response.Update UpdatedResponse
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Content).ToEqual "Ok"
    End With
    
    ' CreateFromHttp
    
    ' CreateFromCURL
    ' --------------------------------------------- '
    With Specs.It("CreateFromCURL should handle 100 Continue responses")
        Set Client = New WebClient
        Set Request = New WebRequest
        Set Response = New WebResponse
        
        Request.Format = WebFormat.PlainText
        Curl = "HTTP/1.1 100 Continue" & web_CrLf & _
            web_CrLf & _
            "HTTP/1.1 100 Continue" & web_CrLf & _
            web_CrLf & _
            "HTTP/1.1 200 OK" & web_CrLf & _
            "Set-Cookie: message=Howdy!" & web_CrLf & _
            web_CrLf & _
            "Text"
        
        Response.CreateFromCurl Client, Request, Curl
        
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Ok
        .Expect(Response.StatusDescription).ToEqual "OK"
        .Expect(Response.Cookies.Count).ToBeGT 0
        .Expect(WebHelpers.FindInKeyValues(Response.Cookies, "message")).ToEqual "Howdy!"
        .Expect(Response.Content).ToEqual "Text"
    End With
    
    ' ExtractHeaders
    ' --------------------------------------------- '
    With Specs.It("ExtractHeaders should extract headers from response headers")
        Set Response = New WebResponse
        ResponseHeaders = "Connection: keep -alive" & web_CrLf & _
            "Date: Tue, 18 Feb 2014 15:00:26 GMT" & web_CrLf & _
            "Content-Length: 2" & web_CrLf & _
            "Content-Type: text/plain" & web_CrLf & _
            "Set-Cookie: cookie=simple-cookie; Path=/"

        Set Headers = Response.ExtractHeaders(ResponseHeaders)
        .Expect(Headers.Count).ToEqual 5
        .Expect(WebHelpers.FindInKeyValues(Headers, "Content-Length")).ToEqual "2"
        .Expect(Headers(5)("Key")).ToEqual "Set-Cookie"
        .Expect(Headers(5)("Value")).ToEqual "cookie=simple-cookie; Path=/"
    End With
    
    With Specs.It("ExtractHeaders should extract multi-line headers from response headers")
        Set Response = New WebResponse
        ResponseHeaders = "Connection: keep-alive" & web_CrLf & _
            "Date: Tue, 18 Feb 2014 15:00:26 GMT" & web_CrLf & _
            "WWW-Authenticate: Digest realm=""abc@host.com""" & web_CrLf & _
            "nonce=""abc""" & web_CrLf & _
            "qop=auth" & web_CrLf & _
            "opaque=""abc""" & web_CrLf & _
            "Set-Cookie: cookie=simple-cookie; Path=/"

        Set Headers = Response.ExtractHeaders(ResponseHeaders)
        .Expect(Headers.Count).ToEqual 4
        .Expect(Headers.Item(3)("Key")).ToEqual "WWW-Authenticate"
        .Expect(Headers.Item(3)("Value")).ToEqual "Digest realm=""abc@host.com""" & web_CrLf & _
            "nonce=""abc""" & web_CrLf & _
            "qop=auth" & web_CrLf & _
            "opaque=""abc"""
    End With
    
    ' ExtractCookies
    ' --------------------------------------------- '
    With Specs.It("should extract cookies from response headers")
        Set Response = New WebResponse
        ResponseHeaders = "Connection: keep -alive" & web_CrLf & _
            "Date: Tue, 18 Feb 2014 15:00:26 GMT" & web_CrLf & _
            "Content-Length: 2" & web_CrLf & _
            "Content-Type: text/plain" & web_CrLf & _
            "Set-Cookie: unsigned-cookie=simple-cookie; Path=/" & web_CrLf & _
            "Set-Cookie: signed-cookie=s%3Aspecial-cookie.1Ghgw2qpDY93QdYjGFPDLAsa3%2FI0FCtO%2FvlxoHkzF%2BY; Path=/" & web_CrLf & _
            "Set-Cookie: duplicate-cookie=A; Path=/" & web_CrLf & _
            "Set-Cookie: duplicate-cookie=B" & web_CrLf & _
            "X-Powered-By: Express"

        Set Headers = Response.ExtractHeaders(ResponseHeaders)
        Set Cookies = Response.ExtractCookies(Headers)
        .Expect(Cookies.Count).ToEqual 4
        .Expect(WebHelpers.FindInKeyValues(Cookies, "unsigned-cookie")).ToEqual "simple-cookie"
    End With
    
    With Specs.It("should use RFC 6265 for decoding cookies")
        Set Response = New WebResponse
        ResponseHeaders = "Set-Cookie: a=plus+plus" & web_CrLf & _
            "Set-Cookie: b=""quotes""" & web_CrLf & _
            "Set-Cookie: c=semi-colon; Path=/"

        Set Headers = Response.ExtractHeaders(ResponseHeaders)
        Set Cookies = Response.ExtractCookies(Headers)

        .Expect(WebHelpers.FindInKeyValues(Cookies, "a")).ToEqual "plus+plus"
        .Expect(WebHelpers.FindInKeyValues(Cookies, "b")).ToEqual "quotes"
        .Expect(WebHelpers.FindInKeyValues(Cookies, "c")).ToEqual "semi-colon"
    End With
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    On Error Resume Next
    
    InlineRunner.RunSuite Specs
End Function
