Attribute VB_Name = "RestClientSpecs"
''
' RestClientSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' General and sync specs for the RestClient class
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "RestClient"
    
    Dim Client As New RestClient
    Dim Request As RestRequest
    Dim Response As RestResponse
    Dim Body As Dictionary
    Dim BodyToString As String
    Dim i As Integer
    Dim Options As Dictionary
    Dim XMLBody As Object
    
    Client.BaseUrl = "http://localhost:3000/"
    
    With Specs.It("should return status code and status description from request")
        Set Request = New RestRequest
        Request.Resource = "status/{code}"
        
        Request.AddUrlSegment "code", 200
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.StatusDescription).ToEqual "OK"
        
        Request.AddUrlSegment "code", 304
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 304
        .Expect(Response.StatusDescription).ToEqual "Not Modified"
        
        Request.AddUrlSegment "code", 404
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 404
        .Expect(Response.StatusDescription).ToEqual "Not Found"
        
        Request.AddUrlSegment "code", 500
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 500
        .Expect(Response.StatusDescription).ToEqual "Internal Server Error"
    End With
    
    With Specs.It("should parse request response")
        Set Request = New RestRequest
        Request.Resource = "json"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Data("a")).ToEqual "1"
        .Expect(Response.Data("b")).ToEqual 2
        .Expect(Response.Data("c")).ToEqual 3.14
        .Expect(Response.Data("d")).ToEqual False
        .Expect(Response.Data("e")(1)).ToEqual 4
        .Expect(Response.Data("f")("b")).ToEqual 2
    End With
    
    With Specs.It("should use headers in request")
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddHeader "custom", "Howdy!"
        Request.ContentType = "text/plain"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Data("headers")("content-type")).ToEqual "text/plain"
        .Expect(Response.Data("headers")("custom")).ToEqual "Howdy!"
    End With
    
    With Specs.It("should use http verb in request")
        Set Request = New RestRequest

        Request.Method = httpPOST
        Request.Resource = "post"
        .Expect(Client.Execute(Request).StatusCode).ToEqual 200
        
        Request.Method = httpPUT
        Request.Resource = "put"
        .Expect(Client.Execute(Request).StatusCode).ToEqual 200
    End With
    
    With Specs.It("should use body in request")
        Set Request = New RestRequest
        Request.Resource = "text"
        Request.Method = httpPOST
        Request.ContentType = "text/plain"
        Request.AddBodyString "Howdy!"
        
        .Expect(Client.Execute(Request).Data("body")).ToEqual "Howdy!"
        
        Set Body = New Dictionary
        Body.Add "a", 3.14
        
        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Method = httpPOST
        Request.AddBody Body
        .Expect(Client.Execute(Request).Data("body")("a")).ToEqual 3.14
    End With
    
    With Specs.It("should pass querystring with request")
        Set Request = New RestRequest
        Request.AddQuerystringParam "a", 1
        Request.AddQuerystringParam "b", 3.14
        Request.AddQuerystringParam "c", "Howdy!"
        Request.AddQuerystringParam "d", False
        Request.Resource = "get"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Data("query")("a")).ToEqual "1"
        .Expect(Response.Data("query")("b")).ToEqual "3.14"
        .Expect(Response.Data("query")("c")).ToEqual "Howdy!"
        .Expect(Response.Data("query")("d")).ToEqual "false"
    End With
    
    With Specs.It("should GET json")
        Set Response = Client.GetJSON("/get")
        
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToBeDefined
    End With
    
    With Specs.It("should POST json")
        Set Body = New Dictionary
        Body.Add "a", 3.14
        Set Response = Client.PostJSON("/post", Body)
        
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("body")("a")).ToEqual 3.14
        
        Set Response = Client.PostJSON("/post", Array(1, 2, 3))
        
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("body")(1)).ToEqual 1
    End With
    
    With Specs.It("should include options with GET and POST json")
        Set Options = New Dictionary
        Options.Add "Headers", New Dictionary
        Options("Headers").Add "custom", "value"
        Set Response = Client.GetJSON("/get", Options)
        
        .Expect(Response.Data("headers")("custom")).ToEqual "value"
    End With
    
    With Specs.It("should return 408 on request timeout")
        Set Request = New RestRequest
        Request.Resource = "timeout"
        Request.AddQuerystringParam "ms", 2000

        Client.TimeoutMS = 500
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 408
        .Expect(Response.StatusDescription).ToEqual "Request Timeout"
        Client.TimeoutMS = 2000
    End With

    With Specs.It("should add content-length header (if enabled)")
        Set Request = New RestRequest
        Request.Resource = "text"
        Request.Method = httpPOST
        Request.ContentType = "text/plain"
        Request.AddBodyString "Howdy!"
        
        Set Response = Client.Execute(Request)
        .Expect(Request.Headers("Content-Length")).ToEqual "6"
        
        Request.IncludeContentLength = False
        Set Response = Client.Execute(Request)
        .Expect(Request.Headers.Exists("Content-Length")).ToEqual False
        
        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Method = httpPOST
        
        Set Body = New Dictionary
        Body.Add "a", 3.14
        Request.AddBody Body
        
        Set Response = Client.Execute(Request)
        .Expect(Request.Headers("Content-Length")).ToEqual "10"
        
        Request.IncludeContentLength = False
        Set Response = Client.Execute(Request)
        .Expect(Request.Headers.Exists("Content-Length")).ToEqual False
    End With
    
    With Specs.It("should include binary body in response")
        Set Request = New RestRequest
        Request.Resource = "howdy"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Body).ToBeDefined
        
        If Not IsEmpty(Response.Body) Then
            For i = LBound(Response.Body) To UBound(Response.Body)
                BodyToString = BodyToString & Chr(Response.Body(i))
            Next i
        End If
        
        .Expect(BodyToString).ToEqual "Howdy!"
    End With
    
    With Specs.It("should include headers in response")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Headers.Count).ToBeGTE 5
        
        Dim Header As Dictionary
        Dim NumCookies As Integer
        For Each Header In Response.Headers
            If Header("key") = "Set-Cookie" Then
                NumCookies = NumCookies + 1
            End If
        Next Header
        
        .Expect(NumCookies).ToEqual 5
    End With

    With Specs.It("should include cookies in response")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Cookies.Count).ToEqual 4
        .Expect(Response.Cookies("unsigned-cookie")).ToEqual "simple-cookie"
        .Expect(Response.Cookies("signed-cookie")).ToContain "special-cookie"
        .Expect(Response.Cookies("tricky;cookie")).ToEqual "includes; semi-colon and space at end "
        .Expect(Response.Cookies("duplicate-cookie")).ToEqual "B"
    End With
    
    With Specs.It("should include cookies with request")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        Set Response = Client.Execute(Request)
    
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddCookie "test-cookie", "howdy"
        Request.AddCookie "signed-cookie", Response.Cookies("signed-cookie")
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Data("cookies").Count).ToEqual 1
        .Expect(Response.Data("cookies")("test-cookie")).ToEqual "howdy"
        .Expect(Response.Data("signed_cookies").Count).ToEqual 1
        .Expect(Response.Data("signed_cookies")("signed-cookie")).ToEqual "special-cookie"
    End With
    
    With Specs.It("should allow separate request and response formats")
        Set Request = New RestRequest
        Request.Resource = "post"
        
        Request.AddParameter "a", 123
        Request.AddParameter "b", 456
        Request.RequestFormat = AvailableFormats.formurlencoded
        Request.ResponseFormat = AvailableFormats.json
        Request.Method = httpPOST
        
        Set Response = Client.Execute(Request)
        
        .Expect(Request.Body).ToEqual "a=123&b=456"
        .Expect(Response.Data("headers")("content-type")).ToEqual "application/x-www-form-urlencoded;charset=UTF-8"
        .Expect(Response.Data("headers")("accept")).ToEqual "application/json"
    End With
    
    With Specs.It("should convert and parse json")
        Set Request = New RestRequest
        Request.Resource = "json"
        Request.Format = json
        Request.Method = httpGET
        
        Set Body = New Dictionary
        Body.Add "a", 123
        Body.Add "b", 456
        Request.AddBody Body
        
        Set Response = Client.Execute(Request)
        
        .Expect(Request.Body).ToEqual "{""a"":123,""b"":456}"
        .Expect(Response.Data("a")).ToEqual "1"
        .Expect(Response.Data("b")).ToEqual 2
        .Expect(Response.Data("c")).ToEqual 3.14
    End With
    
    With Specs.It("should convert and part url-encoded")
        Set Request = New RestRequest
        Request.Resource = "formurlencoded"
        Request.Format = formurlencoded
        Request.Method = httpGET
        
        Set Body = New Dictionary
        Body.Add "a", 123
        Body.Add "b", 456
        Request.AddBody Body
        
        Set Response = Client.Execute(Request)
        
        .Expect(Request.Body).ToEqual "a=123&b=456"
        .Expect(Response.Data("a")).ToEqual "1"
        .Expect(Response.Data("b")).ToEqual "2"
        .Expect(Response.Data("c")).ToEqual "3.14"
    End With
    
    With Specs.It("should convert and parse XML")
        Set Request = New RestRequest
        Request.Resource = "xml"
        Request.Format = xml
        Request.Method = httpGET
        
        Set XMLBody = CreateObject("MSXML2.DOMDocument")
        XMLBody.Async = False
        XMLBody.LoadXML "<Point><X>1.23</X><Y>4.56</Y></Point>"
        Request.AddBody XMLBody

        Set Response = Client.Execute(Request)
    
        .Expect(Request.Body).ToEqual "<Point><X>1.23</X><Y>4.56</Y></Point>"
        .Expect(Response.Content).ToEqual "<Point><X>1.23</X><Y>4.56</Y></Point>"
        .Expect(Response.Data.FirstChild.SelectSingleNode("X").Text).ToEqual "1.23"
        .Expect(Response.Data.FirstChild.SelectSingleNode("Y").Text).ToEqual "4.56"
    End With
    
    With Specs.It("should convert and parse plaintext")
        Set Request = New RestRequest
        Request.Resource = "howdy"
        Request.Format = plaintext
        Request.Method = httpGET
        
        Request.AddBody "Hello?"
        Set Response = Client.Execute(Request)
        
        .Expect(Request.Body).ToEqual "Hello?"
        .Expect(Response.Content).ToEqual "Howdy!"
        .Expect(Response.Data).ToBeUndefined
    End With
    
    With Specs.It("should parse GZIP response")
        Set Request = New RestRequest
        Request.Resource = "json"
        Request.Format = json
        Request.Method = httpGET
        Request.AddHeader "Accept-Encoding", "gzip, deflate"
        
        Set Body = New Dictionary
        Body.Add "a", 123
        Body.Add "b", 456
        Request.AddBody Body
        
        Set Response = Client.Execute(Request)
        
        .Expect(Request.Body).ToEqual "{""a"":123,""b"":456}"
        .Expect(Response.Data("a")).ToEqual "1"
        .Expect(Response.Data("b")).ToEqual 2
        .Expect(Response.Data("c")).ToEqual 3.14
    End With
    
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function

