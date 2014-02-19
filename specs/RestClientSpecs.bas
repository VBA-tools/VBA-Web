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
    
    Client.BaseUrl = "localhost:3000/"
    
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
        .Expect(Response.Data("query")("d")).ToEqual "False"
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

    With Specs.It("should include cookies in response")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Cookies.count).ToEqual 4
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
        .Expect(Response.Data("cookies").count).ToEqual 1
        .Expect(Response.Data("cookies")("test-cookie")).ToEqual "howdy"
        .Expect(Response.Data("signed_cookies").count).ToEqual 1
        .Expect(Response.Data("signed_cookies")("signed-cookie")).ToEqual "special-cookie"
    End With
    
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function

