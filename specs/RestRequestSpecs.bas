Attribute VB_Name = "RestRequestSpecs"
''
' RestRequestSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Specs for the RestRequest class
'
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Dim Request As RestRequest
    
    Specs.Description = "RestRequest"
    
    With Specs.It("should replace url segments for FormattedResource")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False
        
        Request.Resource = "{a1}/{b2}/{c3}/{a1/b2/c3}"
        Request.AddUrlSegment "a1", "A"
        Request.AddUrlSegment "b2", "B"
        Request.AddUrlSegment "c3", "C"
        Request.AddUrlSegment "a1/b2/c3", "D"
        
        .Expect(Request.FormattedResource).ToEqual "A/B/C/D"
    End With
    
    With Specs.It("should only include parameters in querystring for GET requests")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False
        
        Request.AddParameter "A", 123
        
        Request.Method = httpGET
        .Expect(Request.FormattedResource).ToEqual "?A=123"
        
        Request.Method = httpPOST
        .Expect(Request.FormattedResource).ToEqual ""
        Request.Method = httpPUT
        .Expect(Request.FormattedResource).ToEqual ""
        Request.Method = httpPATCH
        .Expect(Request.FormattedResource).ToEqual ""
        Request.Method = httpDELETE
        .Expect(Request.FormattedResource).ToEqual ""
    End With
    
    With Specs.It("should include querystring parameters in FormattedResource for all request types")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False

        Request.AddQuerystringParam "A", 123
        
        Request.Method = httpGET
        .Expect(Request.FormattedResource).ToEqual "?A=123"
        Request.Method = httpPOST
        .Expect(Request.FormattedResource).ToEqual "?A=123"
        Request.Method = httpPUT
        .Expect(Request.FormattedResource).ToEqual "?A=123"
        Request.Method = httpPATCH
        .Expect(Request.FormattedResource).ToEqual "?A=123"
        Request.Method = httpDELETE
        .Expect(Request.FormattedResource).ToEqual "?A=123"
    End With

    With Specs.It("should have ? and add & between parameters for querystring")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False

        Request.AddParameter "A", 123
        Request.AddParameter "B", "456"
        Request.AddQuerystringParam "C", 789
        Request.Method = httpGET
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456&C=789"
    End With

    With Specs.It("should URL encode querystring")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False
    
        Request.AddParameter "A B", " !""#$%&'"
        Request.Method = httpGET
        
        .Expect(Request.FormattedResource).ToEqual "?A%20B=%20%21%22%23%24%25%26%27"
    End With
    
    With Specs.It("should include cachebreaker in FormattedResource by default")
        Set Request = New RestRequest
        
        .Expect(Request.FormattedResource).ToContain "cachebreaker"
    End With
    
    With Specs.It("should be able to remove cachebreaker from FormattedResource")
        Set Request = New RestRequest
        
        Request.IncludeCacheBreaker = False
        .Expect(Request.FormattedResource).ToNotContain "cachebreaker"
    End With
    
    With Specs.It("should add cachebreaker to querystring only")
        Set Request = New RestRequest
        
        Request.Method = httpGET
        .Expect(Request.FormattedResource).ToContain "cachebreaker"
        
        Request.Method = httpPOST
        .Expect(Request.FormattedResource).ToContain "cachebreaker"
        .Expect(Request.Body).ToEqual ""
    End With
    
    With Specs.It("should use body string directly if no parameters")
        Set Request = New RestRequest
        
        Request.AddBodyString "ABC"
        .Expect(Request.Body).ToEqual "ABC"
    End With

    With Specs.It("should only combine body and parameters if not GET Request")
        Set Request = New RestRequest

        Dim Body As New Dictionary
        Body.Add "A", 123
        
        Request.AddBody Body
        Request.AddParameter "b", 456
        
        Request.Method = httpGET
        .Expect(Request.Body).ToEqual "{""A"":123}"
        
        Request.Method = httpPOST
        .Expect(Request.Body).ToEqual "{""A"":123,""b"":456}"
    End With
    
    With Specs.It("should use given client base url for FullUrl only if BaseUrl isn't already set")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False
        Request.RequireHTTPS = True
        
        Request.Resource = "status"
        
        .Expect(Request.FullUrl("facebook.com/api")).ToEqual "https://facebook.com/api/status"
    End With
    
    With Specs.It("should automatically add slash between base and resource")
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False
        Request.RequireHTTPS = True
        
        Request.Resource = "status"
        .Expect(Request.FullUrl("facebook.com/api")).ToEqual "https://facebook.com/api/status"
        
        Request.Resource = "/status"
        .Expect(Request.FullUrl("facebook.com/api")).ToEqual "https://facebook.com/api/status"
        
        Request.Resource = "status"
        .Expect(Request.FullUrl("facebook.com/api/")).ToEqual "https://facebook.com/api/status"
        
        Request.Resource = "/status"
        .Expect(Request.FullUrl("facebook.com/api/")).ToEqual "https://facebook.com/api/status"
    End With
    
    With Specs.It("should user form-urlencoded content type for non-GET requests with parameters")
        Set Request = New RestRequest
        
        Request.AddParameter "A", 123
        Request.Method = httpPOST
        
        .Expect(Request.ContentType).ToEqual "application/x-www-form-urlencoded;charset=UTF-8"
    End With
    
    With Specs.It("should use application/json for GET requests with parameters and requests without parameters")
        Set Request = New RestRequest
        
        Request.Method = httpPOST
        .Expect(Request.ContentType).ToEqual "application/json"
        
        Request.AddParameter "A", 123
        Request.Method = httpGET
        
        .Expect(Request.ContentType).ToEqual "application/json"
    End With
    
    With Specs.It("should override existing headers, url segments, and parameters")
        Set Request = New RestRequest
        
        Request.AddHeader "A", "abc"
        Request.AddHeader "A", "def"
        .Expect(Request.Headers.count).ToEqual 1
        .Expect(Request.Headers.Item("A")).ToEqual "def"
        
        Request.AddUrlSegment "A", "abc"
        Request.AddUrlSegment "A", "def"
        .Expect(Request.UrlSegments.count).ToEqual 1
        .Expect(Request.UrlSegments.Item("A")).ToEqual "def"
        
        Request.AddParameter "A", "abc"
        Request.AddParameter "A", "def"
        .Expect(Request.Parameters.count).ToEqual 1
        .Expect(Request.Parameters.Item("A")).ToEqual "def"
        
        Request.AddQuerystringParam "A", "abc"
        Request.AddQuerystringParam "A", "def"
        .Expect(Request.QuerystringParams.count).ToEqual 1
        .Expect(Request.QuerystringParams.Item("A")).ToEqual "def"
    End With
    
    With Specs.It("should handle Integer, Double, and Boolean variable types as parameters")
        Set Request = New RestRequest
        
        Dim A As Integer
        Dim B As Double
        Dim C As Boolean
        
        A = 20
        B = 3.14
        C = True
        
        Request.AddParameter "A", A
        Request.AddParameter "B", B
        Request.AddParameter "C", C
        
        Request.Method = httpPOST
        .Expect(Request.Body).ToEqual "{""A"":20,""B"":3.14,""C"":true}"
        
        Request.Method = httpGET
        Request.IncludeCacheBreaker = False
        .Expect(Request.FormattedResource).ToEqual "?A=20&B=3.14&C=True"
    End With
    
    InlineRunner.RunSuite Specs
End Function
