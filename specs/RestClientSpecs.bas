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

Public Property Get BaseUrl() As String
    BaseUrl = "http://httpbin.org"
End Property

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
    
    On Error Resume Next
    Client.BaseUrl = BaseUrl
    Client.TimeoutMS = 5000
    
    With Specs.It("should return status code and status description from request")
        Set Request = New RestRequest
        Request.Resource = "status/{code}"
        
        Request.AddUrlSegment "code", 200
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "OK"
        
        Request.AddUrlSegment "code", 304
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 304
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NOT MODIFIED"
        
        Request.AddUrlSegment "code", 404
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 404
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "NOT FOUND"
        
        Request.AddUrlSegment "code", 500
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 500
        .Expect(VBA.UCase$(Response.StatusDescription)).ToEqual "INTERNAL SERVER ERROR"
    End With
    
    With Specs.It("should parse request response")
        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Method = httpPOST
        Request.AddBodyParameter "a", "1"
        Request.AddBodyParameter "b", 2
        Request.AddBodyParameter "c", 3.14
        Request.AddBodyParameter "d", False
        Request.AddBodyParameter "e", Array(1)
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")("a")).ToEqual "1"
        .Expect(Response.Data("json")("b")).ToEqual 2
        .Expect(Response.Data("json")("c")).ToEqual 3.14
        .Expect(Response.Data("json")("d")).ToEqual False
        .Expect(Response.Data("json")("e")(1)).ToEqual 1
    End With

    With Specs.It("should use headers in request")
        Set Request = New RestRequest
        Request.Resource = "headers"
        Request.AddHeader "Custom", "Howdy!"
        Request.ContentType = "text/plain"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("headers")("Content-Type")).ToEqual "text/plain"
        .Expect(Response.Data("headers")("Custom")).ToEqual "Howdy!"
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
        Request.Resource = "post"
        Request.Method = httpPOST
        Request.ContentType = "text/plain"
        Request.Body = "Howdy!"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("data")).ToEqual "Howdy!"

        Set Body = New Dictionary
        Body.Add "a", 3.14

        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Method = httpPOST
        Set Request.Body = Body

        Set Response = Client.Execute(Request)
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")("a")).ToEqual 3.14
    End With

    With Specs.It("should pass querystring with request")
        Set Request = New RestRequest
        Request.AddQuerystringParam "a", 1
        Request.AddQuerystringParam "b", 3.14
        Request.AddQuerystringParam "c", "Howdy!"
        Request.AddQuerystringParam "d", False
        Request.Resource = "get"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("args")("a")).ToEqual "1"
        .Expect(Response.Data("args")("b")).ToEqual "3.14"
        .Expect(Response.Data("args")("c")).ToEqual "Howdy!"
        .Expect(Response.Data("args")("d")).ToEqual "false"
    End With

    With Specs.It("should GET json")
        Set Response = Client.GetJSON("/get")

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
    End With

    With Specs.It("should POST json")
        Set Body = New Dictionary
        Body.Add "a", 3.14
        Set Response = Client.PostJSON("/post", Body)

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")("a")).ToEqual 3.14

        Set Response = Client.PostJSON("/post", Array(1, 2, 3))

        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")(1)).ToEqual 1
    End With

    With Specs.It("should include options with GET and POST json")
        Set Options = New Dictionary
        Options.Add "Headers", New Collection
        Options("Headers").Add RestHelpers.CreateKeyValue("Custom", "value")
        Set Response = Client.GetJSON("/get", Options)

        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("headers")("Custom")).ToEqual "value"
    End With
    
    With Specs.It("should automatically add slash between base and resource")
        Set Request = New RestRequest
    
        Client.BaseUrl = "https://facebook.com/api"
        Request.Resource = "status"
        .Expect(Client.GetFullUrl(Request.FormattedResource)).ToEqual "https://facebook.com/api/status"
    
        Client.BaseUrl = "https://facebook.com/api"
        Request.Resource = "/status"
        .Expect(Client.GetFullUrl(Request.FormattedResource)).ToEqual "https://facebook.com/api/status"
    
        Client.BaseUrl = "https://facebook.com/api/"
        Request.Resource = "status"
        .Expect(Client.GetFullUrl(Request.FormattedResource)).ToEqual "https://facebook.com/api/status"

        Client.BaseUrl = "https://facebook.com/api/"
        Request.Resource = "/status"
        .Expect(Client.GetFullUrl(Request.FormattedResource)).ToEqual "https://facebook.com/api/status"
        
        Client.BaseUrl = BaseUrl
    End With
    
    With Specs.It("should add content-length header")
        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Method = httpPOST
        Request.ContentType = "text/plain"
        Request.Body = "Howdy!"

        Set Response = Client.Execute(Request)
        .Expect(RestHelpers.FindInKeyValues(Request.Headers, "Content-Length")).ToEqual "6"

        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Method = httpPOST

        Set Body = New Dictionary
        Body.Add "a", 3.14
        Set Request.Body = Body

        Set Response = Client.Execute(Request)
        .Expect(RestHelpers.FindInKeyValues(Request.Headers, "Content-Length")).ToEqual "10"
    End With

    With Specs.It("should include binary body in response")
        Set Request = New RestRequest
        Request.Resource = "robots.txt"

        Set Response = Client.Execute(Request)
        .Expect(Response.Body).ToNotBeUndefined

        If Not IsEmpty(Response.Body) Then
            For i = LBound(Response.Body) To UBound(Response.Body)
                BodyToString = BodyToString & Chr(Response.Body(i))
            Next i
        End If

        .Expect(BodyToString).ToEqual "User-agent: *" & vbLf & "Disallow: /deny" & vbLf
    End With

    With Specs.It("should include headers in response")
        Set Request = New RestRequest
        Request.Resource = "response-headers"
        Request.AddQuerystringParam "X-Custom", "Howdy!"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.Headers.Count).ToBeGTE 1
        
        Dim Header As Dictionary
        Dim CustomValue As String
        For Each Header In Response.Headers
            If Header("Key") = "X-Custom" Then
                CustomValue = Header("Value")
            End If
        Next Header
        
        .Expect(CustomValue).ToEqual "Howdy!"
    End With
    
    With Specs.It("should include cookies in response")
        Set Request = New RestRequest
        Request.Resource = "response-headers"
        Request.AddQuerystringParam "Set-Cookie", "a=abc"
        
        ' TODO Possible once duplicate querystrings are allowed
        ' Request.AddQuerystringParam "Set-Cookie", "b=def"
        
        Set Response = Client.Execute(Request)
'        .Expect(Response.Cookies.Count).ToEqual 2
'        .Expect(Response.Cookies("a")).ToEqual "abc"
'        .Expect(Response.Cookies("b")).ToEqual "def"
        .Expect(Response.Cookies.Count).ToEqual 1
        .Expect(Response.Cookies("a")).ToEqual "abc"
    End With

    With Specs.It("should include cookies with request")
        Set Request = New RestRequest
        Request.Resource = "cookies"
        Request.AddCookie "a", "abc"
        Request.AddCookie "b", "def"

        Set Response = Client.Execute(Request)

        Set Request = New RestRequest
        
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("cookies")("a")).ToEqual "abc"
        .Expect(Response.Data("cookies")("b")).ToEqual "def"
    End With

    With Specs.It("should allow separate request and response formats")
        Set Request = New RestRequest
        Request.Resource = "post"

        Request.AddBodyParameter "a", 123
        Request.AddBodyParameter "b", 456
        Request.RequestFormat = WebFormat.formurlencoded
        Request.ResponseFormat = WebFormat.json
        Request.Method = WebMethod.httpPOST

        Set Response = Client.Execute(Request)

        .Expect(Request.Body).ToEqual "a=123&b=456"
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("headers")("Content-Type")).ToEqual "application/x-www-form-urlencoded;charset=UTF-8"
        .Expect(Response.Data("headers")("Accept")).ToEqual "application/json"
    End With

    With Specs.It("should convert and parse json")
        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Format = WebFormat.json
        Request.Method = WebMethod.httpPOST

        Set Body = New Dictionary
        Body.Add "a", "1"
        Body.Add "b", 2
        Body.Add "c", 3.14
        Set Request.Body = Body

        Set Response = Client.Execute(Request)

        .Expect(Request.Body).ToEqual "{""a"":""1"",""b"":2,""c"":3.14}"
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data("json")("a")).ToEqual "1"
        .Expect(Response.Data("json")("b")).ToEqual 2
        .Expect(Response.Data("json")("c")).ToEqual 3.14
    End With

#If Mac Then
#Else
    With Specs.It("should convert and parse XML")
        Set Request = New RestRequest
        Request.Resource = "xml"
        Request.Format = xml
        Request.Method = httpGET

        Set Response = Client.Execute(Request)
        .Expect(Response.Content).ToMatch "<slideshow"
        .Expect(Response.Data).ToNotBeUndefined
        .Expect(Response.Data.FirstChild.SelectSingleNode("slide").SelectSingleNode("title").Text).ToEqual "Wake up to WonderWidgets!"

        Set XMLBody = CreateObject("MSXML2.DOMDocument")
        XMLBody.Async = False
        XMLBody.LoadXML "<Point><X>1.23</X><Y>4.56</Y></Point>"
        Set Request.Body = XMLBody
        .Expect(Request.Body).ToEqual "<Point><X>1.23</X><Y>4.56</Y></Point>"
    End With
#End If

    With Specs.It("should convert and parse plaintext")
        Set Request = New RestRequest
        Request.Resource = "post"
        Request.Format = plaintext
        Request.Method = httpPOST

        Request.Body = "Hello?"
        Set Response = Client.Execute(Request)

        .Expect(Request.Body).ToEqual "Hello?"
        .Expect(Response.Content).ToMatch """data"": ""Hello?"""
        .Expect(Response.Data).ToBeUndefined
    End With

#If Mac Then
    With Specs.It("should prepare cURL request")
        Set Client = New RestClient
        Client.BaseUrl = "http://localhost:3000/"
        Client.Username = "user"
        Client.Password = "password"
        Client.ProxyServer = "proxyserver"
        Client.ProxyBypassList = "proxyserver:80, *.github.com"
        Client.ProxyUsername = "proxyuser"
        Client.ProxyPassword = "proxypassword"
        
        Set Request = New RestRequest
        Request.Resource = "text"
        Request.AddQuerystringParam "type", "message"
        Request.Method = httpPOST
        Request.RequestFormat = WebFormat.plaintext
        Request.ResponseFormat = WebFormat.json
        Request.Body = "Howdy!"
        Request.AddHeader "custom", "Howdy!"
        Request.AddCookie "test-cookie", "howdy"
        
        Dim cURL As String
        
        cURL = Client.PrepareCURL(Request)
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

    With Specs.It("should return 408 on request timeout")
        Set Client = New RestClient
        Client.BaseUrl = BaseUrl
        Client.TimeoutMS = 500
        
        Set Request = New RestRequest
        Request.Resource = "delay/{seconds}"
        Request.AddUrlSegment "seconds", "2"
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 408
        .Expect(Response.StatusDescription).ToEqual "Request Timeout"
    End With
 
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function

