Attribute VB_Name = "RestHelpersSpecs"
''
' RestHelpersSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Specs for RestHelpers
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "RestHelpers"
    
    ' Contents:
    ' 1. Logging
    ' 2. Converters and encoding
    ' 3. Url handling
    ' 4. Object/Dictionary/Collection helpers
    ' 5. Request preparation / handling
    ' 6. Timing
    ' 7. Cryptography
    ' --------------------------------------------- '
    
    Dim json As String
    Dim Parsed As Object
    Dim Obj As Object
    Dim Coll As Collection
    Dim A As Dictionary
    Dim B As Dictionary
    Dim Combined As Dictionary
    Dim Whitelist As Variant
    Dim Filtered As Dictionary
    Dim Encoded As String
    Dim Parts As Dictionary
    Dim ResponseHeaders As String
    Dim Headers As Collection
    Dim Cookies As Dictionary
    Dim Options As Dictionary
    Dim Request As RestRequest
    Dim Response As RestResponse
    Dim UpdatedResponse As RestResponse
    Dim XMLBody As Object
    
    ' ============================================= '
    ' 2. Converters and encoding
    ' ============================================= '
    
    With Specs.It("should parse json")
        json = "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2]}"
        Set Parsed = RestHelpers.ParseJSON(json)
        
        .Expect(Parsed).ToNotBeUndefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed("a")).ToEqual 1
            .Expect(Parsed("b")).ToEqual 3.14
            .Expect(Parsed("c")).ToEqual "Howdy!"
            .Expect(Parsed("d")).ToEqual True
            .Expect(Parsed("e").Count).ToEqual 2
        End If
        
        json = "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""}]"
        Set Parsed = RestHelpers.ParseJSON(json)
        
        .Expect(Parsed).ToNotBeUndefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed(1)).ToEqual 1
            .Expect(Parsed(2)).ToEqual 3.14
            .Expect(Parsed(3)).ToEqual "Howdy!"
            .Expect(Parsed(4)).ToEqual True
            .Expect(Parsed(5).Count).ToEqual 2
            .Expect(Parsed(6)("a")).ToEqual "Howdy!"
        End If
    End With
    
    With Specs.It("should overwrite parsed json for duplicate keys")
        json = "{""a"":1,""a"":2,""a"":3}"
        Set Parsed = RestHelpers.ParseJSON(json)
        
        .Expect(Parsed).ToNotBeUndefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed("a")).ToEqual 3
        End If
    End With
    
    With Specs.It("should parse json numbers")
        json = "{""a"":1,""b"":1.23,""c"":14.6000000000,""d"":14.6e6,""e"":14.6E6,""f"":10000000000000000000000}"
        Set Parsed = RestHelpers.ParseJSON(json)
        
        .Expect(Parsed).ToNotBeUndefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed("a")).ToEqual 1
            .Expect(Parsed("b")).ToEqual 1.23
            .Expect(Parsed("c")).ToEqual 14.6
            .Expect(Parsed("d")).ToEqual 14600000
            .Expect(Parsed("e")).ToEqual 14600000
            .Expect(Parsed("f")).ToEqual 1E+22
        End If
    End With
    
    With Specs.It("should convert to json")
        Set Obj = New Dictionary
        Obj.Add "a", 1
        Obj.Add "b", 3.14
        Obj.Add "c", "Howdy!"
        Obj.Add "d", True
        Obj.Add "e", Array(1, 2)
        Obj.Add "f", Empty
        Obj.Add "g", Null
        
        json = RestHelpers.ConvertToJSON(Obj)
        .Expect(json).ToEqual "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2],""f"":null,""g"":null}"
        
        Set Obj = New Dictionary
        Obj.Add "a", "Howdy!"
        
        Set Coll = New Collection
        Coll.Add 1
        Coll.Add 3.14
        Coll.Add "Howdy!"
        Coll.Add True
        Coll.Add Array(1, 2)
        Coll.Add Obj
        Coll.Add Empty
        Coll.Add Null
        
        json = RestHelpers.ConvertToJSON(Coll)
        .Expect(json).ToEqual "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""},null,null]"
    End With
    
    With Specs.It("should url encode values")
        .Expect(RestHelpers.UrlEncode("$&+,/:;=?@", EncodeUnsafe:=False)).ToEqual "%24%26%2B%2C%2F%3A%3B%3D%3F%40"
        .Expect(RestHelpers.UrlEncode(" ""<>#%{}|\^~[]`")).ToEqual "%20%22%3C%3E%23%25%7B%7D%7C%5C%5E%7E%5B%5D%60"
        .Expect(RestHelpers.UrlEncode("A + B")).ToEqual "A%20%2B%20B"
        .Expect(RestHelpers.UrlEncode("A + B", SpaceAsPlus:=True)).ToEqual "A+%2B+B"
    End With
    
    With Specs.It("should decode url values")
        .Expect(RestHelpers.UrlDecode("+%20%21%22%23%24%25%26%27")).ToEqual "  !""#$%&'"
        .Expect(RestHelpers.UrlDecode("A%20%2B%20B")).ToEqual "A + B"
        .Expect(RestHelpers.UrlDecode("A+%2B+B")).ToEqual "A + B"
    End With
    
    With Specs.It("should encode string to base64")
        .Expect(RestHelpers.Base64Encode("Howdy!")).ToEqual "SG93ZHkh"
    End With
    
    With Specs.It("should combine and convert parameters to url-encoded string")
        Set A = New Dictionary
        Set B = New Dictionary
        
        A.Add "a", 1
        A.Add "b", 3.14
        B.Add "b", 4.14
        B.Add "c", "Howdy!"
        B.Add "d & e", "A + B"
        
        Encoded = RestHelpers.ConvertToUrlEncoded(RestHelpers.CombineObjects(A, B))
        .Expect(Encoded).ToEqual "a=1&b=4.14&c=Howdy!&d+%26+e=A+%2B+B"
    End With
    
    With Specs.It("should parse url-encoded string")
        Set Parsed = RestHelpers.ParseUrlEncoded("a=1&b=3.14&c=Howdy%21&d+%26+e=A+%2B+B")
        
        .Expect(Parsed("a")).ToEqual "1"
        .Expect(Parsed("b")).ToEqual "3.14"
        .Expect(Parsed("c")).ToEqual "Howdy!"
        .Expect(Parsed("d & e")).ToEqual "A + B"
    End With
    
#If Mac Then
#Else
    With Specs.It("should convert to XML")
        Set XMLBody = CreateObject("MSXML2.DOMDocument")
        XMLBody.Async = False
        XMLBody.LoadXML "<Point><X>1.23</X><Y>4.56</Y></Point>"

        Encoded = RestHelpers.ConvertToXML(XMLBody)
        .Expect(Encoded).ToEqual "<Point><X>1.23</X><Y>4.56</Y></Point>"
    End With
    
    With Specs.It("should parse XML")
        Set Parsed = RestHelpers.ParseXML("<Point><X>1.23</X><Y>4.56</Y></Point>")
        
        .Expect(Parsed.FirstChild.SelectSingleNode("X").Text).ToEqual "1.23"
        .Expect(Parsed.FirstChild.SelectSingleNode("Y").Text).ToEqual "4.56"
    End With
#End If
    
    ' ============================================= '
    ' 3. Url handling
    ' ============================================= '
    
    With Specs.It("should join url with /")
        .Expect(RestHelpers.JoinUrl("a", "b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a/", "b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a", "/b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a/", "/b")).ToEqual "a/b"
    End With
    
    With Specs.It("should not join blank urls with /")
        .Expect(RestHelpers.JoinUrl("", "b")).ToEqual "b"
        .Expect(RestHelpers.JoinUrl("a", "")).ToEqual "a"
    End With
    
    With Specs.It("should identify protocols")
        .Expect(RestHelpers.IncludesProtocol("http://testing.com")).ToEqual "http://"
        .Expect(RestHelpers.IncludesProtocol("https://testing.com")).ToEqual "https://"
        .Expect(RestHelpers.IncludesProtocol("ftp://testing.com")).ToEqual "ftp://"
        .Expect(RestHelpers.IncludesProtocol("//testing.com")).ToEqual ""
        .Expect(RestHelpers.IncludesProtocol("testing.com/http://")).ToEqual ""
        .Expect(RestHelpers.IncludesProtocol("http://https://testing.com")).ToEqual "http://"
    End With
    
    With Specs.It("should remove protocols")
        .Expect(RestHelpers.RemoveProtocol("http://testing.com")).ToEqual "testing.com"
        .Expect(RestHelpers.RemoveProtocol("https://testing.com")).ToEqual "testing.com"
        .Expect(RestHelpers.RemoveProtocol("ftp://testing.com")).ToEqual "testing.com"
        .Expect(RestHelpers.RemoveProtocol("htp://testing.com")).ToEqual "testing.com"
        .Expect(RestHelpers.RemoveProtocol("testing.com/http://")).ToEqual "testing.com/http://"
        .Expect(RestHelpers.RemoveProtocol("http://https://testing.com")).ToEqual "https://testing.com"
    End With
    
    With Specs.It("should extract parts from url")
        Set Parts = RestHelpers.UrlParts("https://www.google.com/dir/1/2/search.html?message=Howdy%20World!&other=123#hash")
        
        .Expect(Parts("Protocol")).ToEqual "https"
        .Expect(Parts("Host")).ToEqual "www.google.com"
        .Expect(Parts("Port")).ToEqual "443"
        .Expect(Parts("Path")).ToEqual "/dir/1/2/search.html"
        .Expect(Parts("Querystring")).ToEqual "message=Howdy%20World!&other=123"
        .Expect(Parts("Hash")).ToEqual "hash"
        
        Set Parts = RestHelpers.UrlParts("localhost:3000/dir/1/2/page%202.html?message=Howdy%20World!&other=123#hash")
        
        .Expect(Parts("Protocol")).ToEqual ""
        .Expect(Parts("Host")).ToEqual "localhost"
        .Expect(Parts("Port")).ToEqual "3000"
        .Expect(Parts("Path")).ToEqual "/dir/1/2/page%202.html"
        .Expect(Parts("Querystring")).ToEqual "message=Howdy%20World!&other=123"
        .Expect(Parts("Hash")).ToEqual "hash"
    End With
    
    ' ============================================= '
    ' 4. Object/Dictionary/Collection helpers
    ' ============================================= '
    
    With Specs.It("should combine objects, with overwrite option")
        Set A = New Dictionary
        Set B = New Dictionary
        
        A.Add "a", 1
        A.Add "b", 3.14
        B.Add "b", 4.14
        B.Add "c", "Howdy!"
        
        Set Combined = RestHelpers.CombineObjects(A, B)
        .Expect(Combined("a")).ToEqual 1
        .Expect(Combined("b")).ToEqual 4.14
        .Expect(Combined("c")).ToEqual "Howdy!"
        
        Set Combined = RestHelpers.CombineObjects(A, B, OverwriteOriginal:=False)
        .Expect(Combined("a")).ToEqual 1
        .Expect(Combined("b")).ToEqual 3.14
        .Expect(Combined("c")).ToEqual "Howdy!"
    End With
    
    With Specs.It("should filter object by whitelist")
        Set Obj = New Dictionary
        Obj.Add "a", 1
        Obj.Add "b", 3.14
        Obj.Add "dangerous", "Howdy!"
        
        Whitelist = Array("a", "b")
        
        Set Filtered = RestHelpers.FilterObject(Obj, Whitelist)
        .Expect(Obj.Exists("dangerous")).ToEqual True
        .Expect(Filtered.Exists("a")).ToEqual True
        .Expect(Filtered.Exists("b")).ToEqual True
        .Expect(Filtered.Exists("dangerous")).ToEqual False
    End With
    
    ' ============================================= '
    ' 5. Request preparation / handling
    ' ============================================= '
    
    With Specs.It("should extract headers from response headers")
        ResponseHeaders = "Connection: keep -alive" & vbCrLf & _
            "Date: Tue, 18 Feb 2014 15:00:26 GMT" & vbCrLf & _
            "Content-Length: 2" & vbCrLf & _
            "Content-Type: text/plain" & vbCrLf & _
            "Set-Cookie: unsigned-cookie=simple-cookie; Path=/" & vbCrLf & _
            "Set-Cookie: signed-cookie=s%3Aspecial-cookie.1Ghgw2qpDY93QdYjGFPDLAsa3%2FI0FCtO%2FvlxoHkzF%2BY; Path=/" & vbCrLf & _
            "Set-Cookie: duplicate-cookie=A; Path=/" & vbCrLf & _
            "Set-Cookie: duplicate-cookie=B" & vbCrLf & _
            "X-Powered-By: Express"
            
        Set Headers = RestHelpers.ExtractHeaders(ResponseHeaders)
        .Expect(Headers.Count).ToEqual 9
        .Expect(Headers.Item(5)("key")).ToEqual "Set-Cookie"
        .Expect(Headers.Item(5)("value")).ToEqual "unsigned-cookie=simple-cookie; Path=/"
    End With
    
    With Specs.It("should extract multi-line headers from response headers")
        ResponseHeaders = "Connection: keep -alive" & vbCrLf & _
            "Date: Tue, 18 Feb 2014 15:00:26 GMT" & vbCrLf & _
            "WWW-Authenticate: Digest realm=""abc@host.com""" & vbCrLf & _
            "nonce=""abc""" & vbCrLf & _
            "qop=auth" & vbCrLf & _
            "opaque=""abc""" & vbCrLf & _
            "Set-Cookie: duplicate-cookie=A; Path=/" & vbCrLf & _
            "Set-Cookie: duplicate-cookie=B" & vbCrLf & _
            "X-Powered-By: Express"
            
        Set Headers = RestHelpers.ExtractHeaders(ResponseHeaders)
        .Expect(Headers.Count).ToEqual 6
        .Expect(Headers.Item(3)("key")).ToEqual "WWW-Authenticate"
        .Expect(Headers.Item(3)("value")).ToEqual "Digest realm=""abc@host.com""" & vbCrLf & _
            "nonce=""abc""" & vbCrLf & _
            "qop=auth" & vbCrLf & _
            "opaque=""abc"""
    End With
    
    With Specs.It("should extract cookies from response headers")
        ResponseHeaders = "Connection: keep -alive" & vbCrLf & _
            "Date: Tue, 18 Feb 2014 15:00:26 GMT" & vbCrLf & _
            "Content-Length: 2" & vbCrLf & _
            "Content-Type: text/plain" & vbCrLf & _
            "Set-Cookie: unsigned-cookie=simple-cookie; Path=/" & vbCrLf & _
            "Set-Cookie: signed-cookie=s%3Aspecial-cookie.1Ghgw2qpDY93QdYjGFPDLAsa3%2FI0FCtO%2FvlxoHkzF%2BY; Path=/" & vbCrLf & _
            "Set-Cookie: duplicate-cookie=A; Path=/" & vbCrLf & _
            "Set-Cookie: duplicate-cookie=B" & vbCrLf & _
            "X-Powered-By: Express"
    
        Set Headers = RestHelpers.ExtractHeaders(ResponseHeaders)
        Set Cookies = RestHelpers.ExtractCookies(Headers)
        .Expect(Cookies.Count).ToEqual 3
        .Expect(Cookies("unsigned-cookie")).ToEqual "simple-cookie"
        .Expect(Cookies("duplicate-cookie")).ToEqual "B"
    End With
    
    With Specs.It("should create request from options")
        Set Request = RestHelpers.CreateRequestFromOptions(Nothing)
        .Expect(Request.Headers.Count).ToEqual 0
        
        Set Options = New Dictionary
        Set Request = RestHelpers.CreateRequestFromOptions(Options)
        .Expect(Request.Headers.Count).ToEqual 0
        
        Options.Add "Headers", New Dictionary
        Options("Headers").Add "HeaderKey", "HeaderValue"
        Set Request = RestHelpers.CreateRequestFromOptions(Options)
        .Expect(Request.Headers("HeaderKey")).ToEqual "HeaderValue"
        
        Options.Add "Cookies", New Dictionary
        Options("Cookies").Add "CookieKey", "CookieValue"
        Set Request = RestHelpers.CreateRequestFromOptions(Options)
        .Expect(Request.Cookies("CookieKey")).ToEqual "CookieValue"
        
        Options.Add "QuerystringParams", New Dictionary
        Options("QuerystringParams").Add "QuerystringKey", "QuerystringValue"
        Set Request = RestHelpers.CreateRequestFromOptions(Options)
        .Expect(Request.QuerystringParams("QuerystringKey")).ToEqual "QuerystringValue"
        
        Options.Add "UrlSegments", New Dictionary
        Options("UrlSegments").Add "SegmentKey", "SegmentValue"
        Set Request = RestHelpers.CreateRequestFromOptions(Options)
        .Expect(Request.UrlSegments("SegmentKey")).ToEqual "SegmentValue"
    End With
    
    With Specs.It("should update response")
        Set Response = New RestResponse
        Set UpdatedResponse = New RestResponse
        
        Response.StatusCode = 401
        Response.Body = Array("Unauthorized")
        Response.Content = "Unauthorized"
        
        UpdatedResponse.StatusCode = 200
        UpdatedResponse.Body = Array("Ok")
        UpdatedResponse.Content = "Ok"
        
        RestHelpers.UpdateResponse Response, UpdatedResponse
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Content).ToEqual "Ok"
    End With
    
    ' ============================================= '
    ' 7. Cryptography
    ' ============================================= '
    
    With Specs.It("should create Nonce of specified length")
        .Expect(Len(RestHelpers.CreateNonce)).ToEqual 32
        .Expect(Len(RestHelpers.CreateNonce(20))).ToEqual 20
    End With
    
    With Specs.It("should MD5 hash string")
        .Expect(RestHelpers.MD5("test")).ToEqual "098f6bcd4621d373cade4e832627b4f6"
        .Expect(RestHelpers.MD5("123456789")).ToEqual "25f9e794323b453885f5181f1b624d0b"
        .Expect(RestHelpers.MD5("Mufasa:testrealm@host.com:Circle Of Life")).ToEqual "939e7578ed9e3c518a452acee763bce9"
    End With
    
    InlineRunner.RunSuite Specs
End Function

