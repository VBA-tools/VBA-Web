Attribute VB_Name = "RestRequestSpecs"
''
' RestRequestSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Specs for the RestRequest class
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Dim Request As RestRequest
    Dim Body As Object
    
    Specs.Description = "RestRequest"
    
    With Specs.It("should replace url segments for FormattedResource")
        Set Request = New RestRequest
        
        Request.Resource = "{a1}/{b2}/{c3}/{a1/b2/c3}"
        Request.AddUrlSegment "a1", "A"
        Request.AddUrlSegment "b2", "B"
        Request.AddUrlSegment "c3", "C"
        Request.AddUrlSegment "a1/b2/c3", "D"
        
        .Expect(Request.FormattedResource).ToEqual "A/B/C/D"
    End With
    
    With Specs.It("should include querystring parameters in FormattedResource for all request types")
        Set Request = New RestRequest

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

        Request.AddQuerystringParam "A", 123
        Request.AddQuerystringParam "B", "456"
        Request.AddQuerystringParam "C", 789
        Request.Method = httpGET
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456&C=789"
    End With
    
    With Specs.It("should not add ? if already in resource")
        Set Request = New RestRequest
        
        Request.AddQuerystringParam "B", "456"
        Request.Method = httpGET
        Request.Resource = "?A=123"
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456"
    End With

    With Specs.It("should URL encode querystring")
        Set Request = New RestRequest
    
        Request.AddQuerystringParam "A B", "$&+,/:;=?@"
        Request.Method = httpGET
        
        .Expect(Request.FormattedResource).ToEqual "?A+B=%24%26%2B%2C%2F%3A%3B%3D%3F%40"
    End With
    
    With Specs.It("should use body string directly if no parameters")
        Set Request = New RestRequest
        
        Request.Body = "ABC"
        .Expect(Request.Body).ToEqual "ABC"
    End With

    With Specs.It("should combine body and body parameters if body is Dictionary")
        Set Request = New RestRequest
        
        Set Body = New Dictionary
        Body.Add "A", 123
        
        Set Request.Body = Body
        Request.AddBodyParameter "b", 456
        
        .Expect(Request.Body).ToEqual "{""A"":123,""b"":456}"
    End With
    
    With Specs.It("should include content-type based on specified format")
        Set Request = New RestRequest
        
        ' JSON by default
        .Expect(Request.ContentType).ToEqual "application/json"
        
        Request.Format = WebFormat.json
        .Expect(Request.ContentType).ToEqual "application/json"
        
        Request.Format = formurlencoded
        .Expect(Request.ContentType).ToEqual "application/x-www-form-urlencoded;charset=UTF-8"
    End With
    
    With Specs.It("should handle Integer, Double, and Boolean variable types as parameters")
        Set Request = New RestRequest
        
        Dim A As Integer
        Dim B As Double
        Dim C As Boolean
        
        A = 20
        B = 3.14
        C = True
        
        Request.AddQuerystringParam "A", A
        Request.AddQuerystringParam "B", B
        Request.AddQuerystringParam "C", C
        
        Request.Method = httpGET
        .Expect(Request.FormattedResource).ToEqual "?A=20&B=3.14&C=true"
    End With
    
    With Specs.It("should allow body or body string for GET requests")
        Set Request = New RestRequest
        Request.Method = httpGET
        
        Set Body = New Dictionary
        Body.Add "A", 123
        
        Set Request.Body = Body
        .Expect(Request.Body).ToEqual "{""A"":123}"
        
        Set Request = New RestRequest
        Request.Method = httpGET
        
        Request.Body = "Howdy!"
        .Expect(Request.Body).ToEqual "Howdy!"
    End With
    
    With Specs.It("should format body based on set format")
        Set Request = New RestRequest
        Request.Method = httpPOST
        
        Request.AddBodyParameter "A", 123
        Request.AddBodyParameter "B", "Howdy!"
        
        ' JSON by default
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""Howdy!""}"
        
        Request.Format = WebFormat.json
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""Howdy!""}"
        
        Request.Format = formurlencoded
        .Expect(Request.Body).ToEqual "A=123&B=Howdy!"
    End With
    
    With Specs.It("should allow array/collection for body")
        Set Request = New RestRequest
        
        Set Body = New Collection
        Body.Add "a"
        Body.Add "b"
        Body.Add "c"
        
        Set Request.Body = Body
        .Expect(Request.Body).ToEqual "[""a"",""b"",""c""]"
        
        Request.Body = Array("a", "b", "c")
        .Expect(Request.Body).ToEqual "[""a"",""b"",""c""]"
    End With
    
    With Specs.It("should clone request")
        Set Body = New Dictionary
        Body.Add "key", "value"
    
        Set Request = New RestRequest
        Request.Accept = "text/plain"
        Request.AddCookie "a", "cookie"
        Request.AddHeader "b", "header"
        Request.AddQuerystringParam "d", "querystring"
        Request.AddUrlSegment "e", "segment"
        Request.ContentType = "application/json"
        Request.Method = httpPOST
        Request.RequestFormat = WebFormat.json
        Request.Resource = "resource/"
        Request.ResponseFormat = plaintext
        Set Request.Body = Body
        
        Dim Cloned As RestRequest
        Set Cloned = Request.Clone
        .Expect(Cloned.Accept).ToEqual "text/plain"
        .Expect(RestHelpers.FindInKeyValues(Cloned.Cookies, "a")).ToEqual "cookie"
        .Expect(RestHelpers.FindInKeyValues(Cloned.Headers, "b")).ToEqual "header"
        .Expect(RestHelpers.FindInKeyValues(Cloned.QuerystringParams, "d")).ToEqual "querystring"
        .Expect(Cloned.UrlSegments("e")).ToEqual "segment"
        .Expect(Cloned.ContentType).ToEqual "application/json"
        .Expect(Cloned.Method).ToEqual httpPOST
        .Expect(Cloned.RequestFormat).ToEqual WebFormat.json
        .Expect(Cloned.Resource).ToEqual "resource/"
        .Expect(Cloned.ResponseFormat).ToEqual plaintext
        .Expect(Cloned.Body).ToEqual "{""key"":""value""}"
        
        Request.Accept = "application/json"
        Request.AddHeader "new", "new_header"
        Request.ResponseFormat = xml
        .Expect(Cloned.Accept).ToEqual "text/plain"
        .Expect(RestHelpers.FindInKeyValues(Cloned.Headers, "new")).ToEqual False
        .Expect(Cloned.ResponseFormat).ToEqual plaintext
    End With
    
    With Specs.It("should have an id")
        Set Request = New RestRequest
        Debug.Print
        .Expect(Len(Request.Id)).ToBeGreaterThan 0
    End With
    
    InlineRunner.RunSuite Specs
End Function


