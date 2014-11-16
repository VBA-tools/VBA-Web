Attribute VB_Name = "Specs_RestRequest"
''
' Specs_RestRequest
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Specs for RestRequest
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    
    Dim Request As RestRequest
    Dim Body As Object
    Dim Cloned As RestRequest
    
    RestHelpers.RegisterConverter "csv", "text/csv", "Specs_RestHelpers.SimpleConverter", "Specs_RestHelpers.SimpleParser"
    
    Specs.Description = "RestRequest"
    
    ' --------------------------------------------- '
    ' Properties
    ' --------------------------------------------- '
    
    ' Resource
    ' Method
    
    ' Body
    ' --------------------------------------------- '
    With Specs.It("Body should Let directly to string")
        Set Request = New RestRequest
        
        Request.Body = "Howdy!"
        Request.Format = WebFormat.plaintext
        
        .Expect(Request.Body).ToEqual "Howdy!"
    End With
    
    With Specs.It("Body should allow Let as Array or Set as Collection")
        Set Request = New RestRequest
        
        Request.Body = Array("A", "B", "C")
        .Expect(Request.Body).ToEqual "[""A"",""B"",""C""]"
        
        Set Body = New Collection
        Body.Add "A"
        Body.Add "B"
        Body.Add "C"
        
        Set Request.Body = Body
        .Expect(Request.Body).ToEqual "[""A"",""B"",""C""]"
    End With
    
    With Specs.It("Body should allow Set as Dictionary")
        Set Request = New RestRequest
        
        Set Body = New Dictionary
        Body.Add "A", 123
        Body.Add "B", "456"
        Body.Add "C", 789
        
        Set Request.Body = Body
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""456"",""C"":789}"
    End With
    
    With Specs.It("Body should be formatted by ResponseFormat")
        Set Request = New RestRequest
        
        Request.AddBodyParameter "A", 123
        Request.AddBodyParameter "B", "Howdy!"
        
        ' JSON by default
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""Howdy!""}"
        
        Request.Format = WebFormat.json
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""Howdy!""}"
        
        Request.Format = WebFormat.formurlencoded
        .Expect(Request.Body).ToEqual "A=123&B=Howdy!"
    End With
    
    With Specs.It("Body should be formatted by CustomRequestFormat")
        Set Request = New RestRequest
        
        Request.AddBodyParameter "message", "Howdy!"
        Request.CustomRequestFormat = "csv"
        
        .Expect(Request.Body).ToEqual "{""message"":""Howdy!"",""response"":""Goodbye!""}"
    End With
    
    ' Format
    ' --------------------------------------------- '
    With Specs.It("Format should set RequestFormat and ResponseFormat")
        Set Request = New RestRequest
        
        Request.Format = WebFormat.plaintext
        
        .Expect(Request.RequestFormat).ToEqual WebFormat.plaintext
        .Expect(Request.ResponseFormat).ToEqual WebFormat.plaintext
    End With
    
    ' RequestFormat
    ' ResponseFormat
    
    ' CustomRequestFormat
    ' --------------------------------------------- '
    With Specs.It("CustomRequestFormat should set RequestFormat to Custom")
        Set Request = New RestRequest
        
        Request.CustomRequestFormat = "csv"
        
        .Expect(Request.RequestFormat).ToEqual WebFormat.custom
    End With
    
    ' CustomResponseFormat
    ' --------------------------------------------- '
    With Specs.It("CustomResponseFormat should set ResponseFormat")
        Set Request = New RestRequest
        
        Request.CustomResponseFormat = "csv"
        
        .Expect(Request.ResponseFormat).ToEqual WebFormat.custom
    End With
    
    ' ContentType
    ' --------------------------------------------- '
    With Specs.It("ContentType should be set from RequestFormat")
        Set Request = New RestRequest
        
        ' JSON by default
        .Expect(Request.ContentType).ToEqual RestHelpers.FormatToMediaType(WebFormat.json)
        
        Request.RequestFormat = WebFormat.json
        .Expect(Request.ContentType).ToEqual RestHelpers.FormatToMediaType(WebFormat.json)
        
        Request.RequestFormat = WebFormat.formurlencoded
        .Expect(Request.ContentType).ToEqual RestHelpers.FormatToMediaType(WebFormat.formurlencoded)
    End With
    
    With Specs.It("ContentType should be set from CustomRequestFormat")
        Set Request = New RestRequest
        
        Request.CustomRequestFormat = "csv"
        
        .Expect(Request.ContentType).ToEqual "text/csv"
    End With
    
    With Specs.It("ContentType should allow override")
        Set Request = New RestRequest
        
        Request.RequestFormat = WebFormat.plaintext
        Request.ContentType = "x-custom/text"
        
        .Expect(Request.ContentType).ToEqual "x-custom/text"
    End With
    
    ' Accept
    ' --------------------------------------------- '
    With Specs.It("Accept should be set from ResponseFormat")
        Set Request = New RestRequest
        
        ' JSON by default
        .Expect(Request.Accept).ToEqual RestHelpers.FormatToMediaType(WebFormat.json)
        
        Request.ResponseFormat = WebFormat.json
        .Expect(Request.Accept).ToEqual RestHelpers.FormatToMediaType(WebFormat.json)
        
        Request.ResponseFormat = WebFormat.formurlencoded
        .Expect(Request.Accept).ToEqual RestHelpers.FormatToMediaType(WebFormat.formurlencoded)
    End With
    
    With Specs.It("Accept should be set from CustomResponseFormat")
        Set Request = New RestRequest
        
        Request.CustomResponseFormat = "csv"
        
        .Expect(Request.Accept).ToEqual "text/csv"
    End With
    
    With Specs.It("Accept should allow override")
        Set Request = New RestRequest
        
        Request.ResponseFormat = WebFormat.plaintext
        Request.Accept = "x-custom/text"
        
        .Expect(Request.Accept).ToEqual "x-custom/text"
    End With
    
    ' ContentLength
    ' --------------------------------------------- '
    With Specs.It("ContentLength should be set from length of Body")
        Set Request = New RestRequest
        
        .Expect(Request.ContentLength).ToEqual 0
        
        Request.Body = "123456789"
        
        .Expect(Request.ContentLength).ToEqual 9
    End With
    
    With Specs.It("ContentLength should allow override")
        Set Request = New RestRequest
        
        Request.Body = "123456789"
        Request.ContentLength = 4
        
        .Expect(Request.ContentLength).ToEqual 4
    End With
    
    ' FormattedResource
    ' --------------------------------------------- '
    With Specs.It("FormattedResource should replace Url Segments")
        Set Request = New RestRequest
        
        Request.Resource = "{a1}/{b2}/{c3}/{a1/b2/c3}"
        Request.AddUrlSegment "a1", "A"
        Request.AddUrlSegment "b2", "B"
        Request.AddUrlSegment "c3", "C"
        Request.AddUrlSegment "a1/b2/c3", "D"
        
        .Expect(Request.FormattedResource).ToEqual "A/B/C/D"
    End With
    
    With Specs.It("FormattedResource should include querystring parameters")
        Set Request = New RestRequest
    
        Request.Resource = "resource"
        Request.AddQuerystringParam "A", 123
        Request.AddQuerystringParam "B", "456"
        Request.AddQuerystringParam "C", 789
        
        .Expect(Request.FormattedResource).ToEqual "resource?A=123&B=456&C=789"
    End With
    
    With Specs.It("FormattedResource should have ? and add & between parameters for querystring")
        Set Request = New RestRequest

        Request.AddQuerystringParam "A", 123
        Request.AddQuerystringParam "B", "456"
        Request.AddQuerystringParam "C", 789
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456&C=789"
    End With
    
    With Specs.It("FormattedResource should not add ? if already in Resource")
        Set Request = New RestRequest
        
        Request.Resource = "?A=123"
        Request.AddQuerystringParam "B", "456"
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456"
    End With

    With Specs.It("FormattedResource should URL encode querystring")
        Set Request = New RestRequest
    
        Request.AddQuerystringParam "A B", "$&+,/:;=?@"
        
        .Expect(Request.FormattedResource).ToEqual "?A+B=%24%26%2B%2C%2F%3A%3B%3D%3F%40"
    End With
    
    ' Cookies
    ' Headers
    ' QuerystringParams
    ' UrlSegments
    
    ' Id
    ' --------------------------------------------- '
    With Specs.It("should have an Id")
        Set Request = New RestRequest
        
        .Expect(Request.Id).ToNotBeUndefined
    End With
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ' AddBodyParameter
    ' --------------------------------------------- '
    With Specs.It("AddBodyParameter should add if Body is Empty")
        Set Request = New RestRequest
        
        Request.AddBodyParameter "A", 123
        Request.AddBodyParameter "B", 456
        
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":456}"
    End With
    
    With Specs.It("AddBodyParameter should add to existing Body if Dictionary")
        Set Request = New RestRequest
        
        Set Body = New Dictionary
        Body.Add "A", 123
        
        Set Request.Body = Body
        Request.AddBodyParameter "B", 456
        
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":456}"
    End With
    
    With Specs.It("AddBodyParameter should throw TODO if adding to existing Body this is not Dictionary")
        On Error Resume Next
        Set Request = New RestRequest
        
        Request.Body = Array("A", "B", "C")
        Request.AddBodyParameter "D", 123
        
        ' TODO Check actual error number
        .Expect(Err.Number).ToNotEqual 0
        .Expect(Err.Description).ToEqual _
            "The existing body is not a Dictionary. Adding body parameters can only be used with Dictionaries"
        
        Err.Clear
        On Error GoTo 0
    End With
    
    ' AddCookie
    ' --------------------------------------------- '
    With Specs.It("should AddCookie")
        Set Request = New RestRequest
        
        Request.AddCookie "A", "cookie"
        Request.AddCookie "B", "cookie 2"
        
        .Expect(Request.Cookies.Count).ToEqual 2
        .Expect(Request.Cookies(1)("Key")).ToEqual "A"
        .Expect(Request.Cookies(2)("Value")).ToEqual "cookie 2"
    End With
    
    ' AddHeader
    ' --------------------------------------------- '
    With Specs.It("should AddHeader")
        Set Request = New RestRequest
        
        Request.AddHeader "A", "header"
        Request.AddHeader "B", "header 2"
        
        .Expect(Request.Headers.Count).ToEqual 2
        .Expect(Request.Headers(1)("Key")).ToEqual "A"
        .Expect(Request.Headers(2)("Value")).ToEqual "header 2"
    End With
    
    ' AddQuerystringParam
    ' --------------------------------------------- '
    With Specs.It("should AddQuerystringParam")
        Set Request = New RestRequest
        
        Request.AddQuerystringParam "A", "querystring"
        Request.AddQuerystringParam "B", "querystring 2"
        
        .Expect(Request.QuerystringParams.Count).ToEqual 2
        .Expect(Request.QuerystringParams(1)("Key")).ToEqual "A"
        .Expect(Request.QuerystringParams(2)("Value")).ToEqual "querystring 2"
    End With
    
    With Specs.It("AddQuerystringParam should allow Integer, Double, and Boolean types")
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
        
        .Expect(Request.FormattedResource).ToEqual "?A=20&B=3.14&C=true"
    End With
    
    With Specs.It("AddQuerystringParam should allow duplicate keys")
        Set Request = New RestRequest
        
        Request.AddQuerystringParam "A", "querystring"
        Request.AddQuerystringParam "A", "querystring 2"
        Request.AddQuerystringParam "A", "querystring 3"
        
        .Expect(Request.QuerystringParams.Count).ToEqual 3
        .Expect(Request.QuerystringParams(1)("Key")).ToEqual "A"
        .Expect(Request.QuerystringParams(2)("Value")).ToEqual "querystring 2"
        .Expect(Request.QuerystringParams(3)("Key")).ToEqual "A"
    End With
    
    ' AddUrlSegment
    ' --------------------------------------------- '
    With Specs.It("should AddUrlSegment")
        Set Request = New RestRequest
        
        Request.AddUrlSegment "A", "segment"
        Request.AddUrlSegment "B", "segment 2"
        
        .Expect(Request.UrlSegments.Count).ToEqual 2
        .Expect(Request.UrlSegments("A")).ToEqual "segment"
        .Expect(Request.UrlSegments("B")).ToEqual "segment 2"
    End With
    
    ' Clone
    ' @internal
    ' --------------------------------------------- '
    With Specs.It("should Clone request, but create unique Id")
        Set Body = New Dictionary
        Body.Add "Key", "Value"
    
        Set Request = New RestRequest
        Request.Resource = "resource/"
        Request.Method = WebMethod.httpPOST
        Request.RequestFormat = WebFormat.json
        Request.ResponseFormat = WebFormat.plaintext
        Request.ContentType = "application/json"
        Request.Accept = "text/plain"
        Request.AddCookie "A", "cookie"
        Request.AddHeader "B", "header"
        Request.AddQuerystringParam "C", "querystring"
        Request.AddUrlSegment "D", "segment"
        Set Request.Body = Body
        
        Set Cloned = Request.Clone
        
        .Expect(Cloned.Resource).ToEqual "resource/"
        .Expect(Cloned.Method).ToEqual WebMethod.httpPOST
        .Expect(Cloned.RequestFormat).ToEqual WebFormat.json
        .Expect(Cloned.ResponseFormat).ToEqual plaintext
        .Expect(Cloned.ContentType).ToEqual "application/json"
        .Expect(Cloned.Accept).ToEqual "text/plain"
    End With
    
    With Specs.It("Clone should have minimal/no references to original")
        Set Body = New Dictionary
        Body.Add "Key", "Value"
    
        Set Request = New RestRequest
        Request.Resource = "resource/"
        Request.Method = WebMethod.httpPOST
        Request.AddCookie "A", "cookie"
        Request.AddHeader "B", "header"
        Request.AddQuerystringParam "C", "querystring"
        Request.AddUrlSegment "D", "segment"
        Set Request.Body = Body
        
        Set Cloned = Request.Clone
        
        Cloned.Resource = "updated/"
        Cloned.Method = WebMethod.httpPUT
        Cloned.AddCookie "E", "cookie"
        Cloned.AddHeader "F", "header"
        Cloned.AddQuerystringParam "G", "querystring"
        Cloned.AddUrlSegment "H", "segment"
        
        .Expect(Request.Resource).ToEqual "resource/"
        .Expect(Request.Method).ToEqual WebMethod.httpPOST
        .Expect(RestHelpers.FindInKeyValues(Request.Cookies, "E")).ToBeUndefined
        .Expect(RestHelpers.FindInKeyValues(Request.Headers, "F")).ToBeUndefined
        .Expect(RestHelpers.FindInKeyValues(Request.QuerystringParams, "G")).ToBeUndefined
        .Expect(Request.UrlSegments.Exists("H")).ToEqual False
    End With
    
    ' Prepare
    ' @internal
    ' --------------------------------------------- '
    With Specs.It("Prepare should add ContentType, Accept, and ContentLength headers")
        Set Request = New RestRequest
        
        Request.ContentType = "text/plain"
        Request.Accept = "text/csv"
        Request.ContentLength = 100
        
        .Expect(Request.Headers.Count).ToEqual 0
        
        Request.Prepare
        
        .Expect(Request.Headers.Count).ToBeGTE 3
        .Expect(RestHelpers.FindInKeyValues(Request.Headers, "Content-Type")).ToEqual "text/plain"
        .Expect(RestHelpers.FindInKeyValues(Request.Headers, "Accept")).ToEqual "text/csv"
        .Expect(RestHelpers.FindInKeyValues(Request.Headers, "Content-Length")).ToEqual "100"
    End With
    
    InlineRunner.RunSuite Specs
End Function


