Attribute VB_Name = "Specs_WebRequest"
''
' Specs_WebRequest
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for WebRequest
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebRequest"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs
    
    Dim Request As WebRequest
    Dim Body As Object
    Dim Cloned As WebRequest
    Dim NumHeaders As Long
    
    WebHelpers.RegisterConverter "csv", "text/csv", "Specs_WebHelpers.SimpleConverter", "Specs_WebHelpers.SimpleParser"
    
    ' --------------------------------------------- '
    ' Properties
    ' --------------------------------------------- '
    
    ' Resource
    ' Method
    
    ' Body
    ' --------------------------------------------- '
    With Specs.It("Body should Let directly to string")
        Set Request = New WebRequest
        
        Request.Body = "Howdy!"
        Request.Format = WebFormat.PlainText
        
        .Expect(Request.Body).ToEqual "Howdy!"
    End With
    
    With Specs.It("Body should allow Let as Array or Set as Collection")
        Set Request = New WebRequest
        
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
        Set Request = New WebRequest
        
        Set Body = New Dictionary
        Body.Add "A", 123
        Body.Add "B", "456"
        Body.Add "C", 789
        
        Set Request.Body = Body
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""456"",""C"":789}"
    End With
    
    With Specs.It("Body should be formatted by ResponseFormat")
        Set Request = New WebRequest
        
        Request.AddBodyParameter "A", 123
        Request.AddBodyParameter "B", "Howdy!"
        
        ' JSON by default
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""Howdy!""}"
        
        Request.Format = WebFormat.Json
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":""Howdy!""}"
        
        Request.Format = WebFormat.FormUrlEncoded
        .Expect(Request.Body).ToEqual "A=123&B=Howdy%21"
    End With
    
    With Specs.It("Body should be formatted by CustomRequestFormat")
        Set Request = New WebRequest
        
        Request.AddBodyParameter "message", "Howdy!"
        Request.CustomRequestFormat = "csv"
        
        .Expect(Request.Body).ToEqual "{""message"":""Howdy!"",""response"":""Goodbye!""}"
    End With
    
    ' Format
    ' --------------------------------------------- '
    With Specs.It("Format should set RequestFormat and ResponseFormat")
        Set Request = New WebRequest
        
        Request.Format = WebFormat.PlainText
        
        .Expect(Request.RequestFormat).ToEqual WebFormat.PlainText
        .Expect(Request.ResponseFormat).ToEqual WebFormat.PlainText
    End With
    
    ' RequestFormat
    ' ResponseFormat
    
    ' CustomRequestFormat
    ' --------------------------------------------- '
    With Specs.It("CustomRequestFormat should set RequestFormat to Custom")
        Set Request = New WebRequest
        
        Request.CustomRequestFormat = "csv"
        
        .Expect(Request.RequestFormat).ToEqual WebFormat.Custom
    End With
    
    ' CustomResponseFormat
    ' --------------------------------------------- '
    With Specs.It("CustomResponseFormat should set ResponseFormat")
        Set Request = New WebRequest
        
        Request.CustomResponseFormat = "csv"
        
        .Expect(Request.ResponseFormat).ToEqual WebFormat.Custom
    End With
    
    ' ContentType
    ' --------------------------------------------- '
    With Specs.It("ContentType should be set from RequestFormat")
        Set Request = New WebRequest
        
        ' JSON by default
        .Expect(Request.ContentType).ToEqual WebHelpers.FormatToMediaType(WebFormat.Json)
        
        Request.RequestFormat = WebFormat.Json
        .Expect(Request.ContentType).ToEqual WebHelpers.FormatToMediaType(WebFormat.Json)
        
        Request.RequestFormat = WebFormat.FormUrlEncoded
        .Expect(Request.ContentType).ToEqual WebHelpers.FormatToMediaType(WebFormat.FormUrlEncoded)
    End With
    
    With Specs.It("ContentType should be set from CustomRequestFormat")
        Set Request = New WebRequest
        
        Request.CustomRequestFormat = "csv"
        
        .Expect(Request.ContentType).ToEqual "text/csv"
    End With
    
    With Specs.It("ContentType should allow override")
        Set Request = New WebRequest
        
        Request.RequestFormat = WebFormat.PlainText
        Request.ContentType = "x-custom/text"
        
        .Expect(Request.ContentType).ToEqual "x-custom/text"
    End With
    
    ' Accept
    ' --------------------------------------------- '
    With Specs.It("Accept should be set from ResponseFormat")
        Set Request = New WebRequest
        
        ' JSON by default
        .Expect(Request.Accept).ToEqual WebHelpers.FormatToMediaType(WebFormat.Json)
        
        Request.ResponseFormat = WebFormat.Json
        .Expect(Request.Accept).ToEqual WebHelpers.FormatToMediaType(WebFormat.Json)
        
        Request.ResponseFormat = WebFormat.FormUrlEncoded
        .Expect(Request.Accept).ToEqual WebHelpers.FormatToMediaType(WebFormat.FormUrlEncoded)
    End With
    
    With Specs.It("Accept should be set from CustomResponseFormat")
        Set Request = New WebRequest
        
        Request.CustomResponseFormat = "csv"
        
        .Expect(Request.Accept).ToEqual "text/csv"
    End With
    
    With Specs.It("Accept should allow override")
        Set Request = New WebRequest
        
        Request.ResponseFormat = WebFormat.PlainText
        Request.Accept = "x-custom/text"
        
        .Expect(Request.Accept).ToEqual "x-custom/text"
    End With
    
    ' ContentLength
    ' --------------------------------------------- '
    With Specs.It("ContentLength should be set from length of Body")
        Set Request = New WebRequest
        
        .Expect(Request.ContentLength).ToEqual 0
        
        Request.Body = "123456789"
        
        .Expect(Request.ContentLength).ToEqual 9
    End With
    
    With Specs.It("ContentLength should allow override")
        Set Request = New WebRequest
        
        Request.Body = "123456789"
        Request.ContentLength = 4
        
        .Expect(Request.ContentLength).ToEqual 4
    End With
    
    ' FormattedResource
    ' --------------------------------------------- '
    With Specs.It("FormattedResource should replace Url Segments")
        Set Request = New WebRequest
        
        Request.Resource = "{a1}/{b2}/{c3}/{a1/b2/c3}"
        Request.AddUrlSegment "a1", "A"
        Request.AddUrlSegment "b2", "B"
        Request.AddUrlSegment "c3", "C"
        Request.AddUrlSegment "a1/b2/c3", "D"
        
        .Expect(Request.FormattedResource).ToEqual "A/B/C/D"
    End With
    
    With Specs.It("FormattedResource should url-encode Url Segments")
        Set Request = New WebRequest
        
        Request.Resource = "{segment}"
        Request.AddUrlSegment "segment", "&/:;=?@"
        
        .Expect(Request.FormattedResource).ToEqual "%26%2F%3A%3B%3D%3F%40"
    End With
    
    With Specs.It("FormattedResource should include querystring parameters")
        Set Request = New WebRequest
    
        Request.Resource = "resource"
        Request.AddQuerystringParam "A", 123
        Request.AddQuerystringParam "B", "456"
        Request.AddQuerystringParam "C", 789
        
        .Expect(Request.FormattedResource).ToEqual "resource?A=123&B=456&C=789"
    End With
    
    With Specs.It("FormattedResource should have ? and add & between parameters for querystring")
        Set Request = New WebRequest

        Request.AddQuerystringParam "A", 123
        Request.AddQuerystringParam "B", "456"
        Request.AddQuerystringParam "C", 789
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456&C=789"
    End With
    
    With Specs.It("FormattedResource should not add ? if already in Resource")
        Set Request = New WebRequest
        
        Request.Resource = "?A=123"
        Request.AddQuerystringParam "B", "456"
        
        .Expect(Request.FormattedResource).ToEqual "?A=123&B=456"
    End With

    With Specs.It("FormattedResource should URL encode querystring with QueryUrlEncoding for non-form-urlencoded")
        Set Request = New WebRequest
    
        Request.AddQuerystringParam "A + B", "*~"
        
        .Expect(Request.FormattedResource).ToEqual "?A%20%2B%20B=%2A%7E"
    End With
    
    With Specs.It("FormattedResource should URL encode querystring with FormUrlEncoding for form-urlencoded")
        Set Request = New WebRequest
        
        Request.RequestFormat = WebFormat.FormUrlEncoded
        Request.AddQuerystringParam "A + B", "*~"
        
        .Expect(Request.FormattedResource).ToEqual "?A+%2B+B=*%7E"
    End With
    
    ' UserAgent
    ' Cookies
    ' Headers
    ' QuerystringParams
    ' UrlSegments
    
    ' Id
    ' --------------------------------------------- '
    With Specs.It("should have an Id")
        Set Request = New WebRequest
        
        .Expect(Request.Id).ToNotBeUndefined
    End With
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ' AddBodyParameter
    ' --------------------------------------------- '
    With Specs.It("AddBodyParameter should add if Body is Empty")
        Set Request = New WebRequest
        
        Request.AddBodyParameter "A", 123
        Request.AddBodyParameter "B", 456
        
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":456}"
    End With
    
    With Specs.It("AddBodyParameter should add to existing Body if Dictionary")
        Set Request = New WebRequest
        
        Set Body = New Dictionary
        Body.Add "A", 123
        
        Set Request.Body = Body
        Request.AddBodyParameter "B", 456
        
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":456}"
    End With
    
    With Specs.It("AddBodyParameter should override cached Body")
        Set Request = New WebRequest
        
        Request.AddBodyParameter "A", 123
        
        .Expect(Request.Body).ToEqual "{""A"":123}"
        
        Request.AddBodyParameter "B", 456
        
        .Expect(Request.Body).ToEqual "{""A"":123,""B"":456}"
    End With
    
    ' AddCookie
    ' --------------------------------------------- '
    With Specs.It("should AddCookie")
        Set Request = New WebRequest
        
        Request.AddCookie "A[1]", "cookie"
        Request.AddCookie "B", "cookie 2"
        
        .Expect(Request.Cookies.Count).ToEqual 2
        .Expect(Request.Cookies(1)("Key")).ToEqual "A%5B1%5D"
        .Expect(Request.Cookies(2)("Value")).ToEqual "cookie%202"
    End With
    
    ' AddHeader
    ' --------------------------------------------- '
    With Specs.It("should AddHeader")
        Set Request = New WebRequest
        
        Request.AddHeader "A", "header"
        Request.AddHeader "B", "header 2"
        
        .Expect(Request.Headers.Count).ToEqual 2
        .Expect(Request.Headers(1)("Key")).ToEqual "A"
        .Expect(Request.Headers(2)("Value")).ToEqual "header 2"
    End With
    
    ' SetHeader
    ' --------------------------------------------- '
    With Specs.It("should SetHeader")
        Set Request = New WebRequest
        
        Request.AddHeader "A", "add"
        
        Request.SetHeader "A", "set"
        Request.SetHeader "B", "header"
        
        .Expect(Request.Headers.Count).ToEqual 2
        .Expect(Request.Headers(1)("Value")).ToEqual "set"
        .Expect(Request.Headers(2)("Key")).ToEqual "B"
    End With
    
    ' AddQuerystringParam
    ' --------------------------------------------- '
    With Specs.It("should AddQuerystringParam")
        Set Request = New WebRequest
        
        Request.AddQuerystringParam "A", "querystring"
        Request.AddQuerystringParam "B", "querystring 2"
        
        .Expect(Request.QuerystringParams.Count).ToEqual 2
        .Expect(Request.QuerystringParams(1)("Key")).ToEqual "A"
        .Expect(Request.QuerystringParams(2)("Value")).ToEqual "querystring 2"
    End With
    
    With Specs.It("AddQuerystringParam should allow Integer, Double, and Boolean types")
        Set Request = New WebRequest
        
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
        Set Request = New WebRequest
        
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
        Set Request = New WebRequest
        
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
    
        Set Request = New WebRequest
        Request.Resource = "resource/"
        Request.Method = WebMethod.HttpPost
        Request.RequestFormat = WebFormat.Json
        Request.ResponseFormat = WebFormat.PlainText
        Request.ContentType = "application/json"
        Request.Accept = "text/plain"
        Request.AddCookie "A", "cookie"
        Request.AddHeader "B", "header"
        Request.AddQuerystringParam "C", "querystring"
        Request.AddUrlSegment "D", "segment"
        Set Request.Body = Body
        
        Set Cloned = Request.Clone
        
        .Expect(Cloned.Resource).ToEqual "resource/"
        .Expect(Cloned.Method).ToEqual WebMethod.HttpPost
        .Expect(Cloned.RequestFormat).ToEqual WebFormat.Json
        .Expect(Cloned.ResponseFormat).ToEqual PlainText
        .Expect(Cloned.ContentType).ToEqual "application/json"
        .Expect(Cloned.Accept).ToEqual "text/plain"
    End With
    
    With Specs.It("Clone should have minimal/no references to original")
        Set Body = New Dictionary
        Body.Add "Key", "Value"
    
        Set Request = New WebRequest
        Request.Resource = "resource/"
        Request.Method = WebMethod.HttpPost
        Request.AddCookie "A", "cookie"
        Request.AddHeader "B", "header"
        Request.AddQuerystringParam "C", "querystring"
        Request.AddUrlSegment "D", "segment"
        Set Request.Body = Body
        
        Set Cloned = Request.Clone
        
        Cloned.Resource = "updated/"
        Cloned.Method = WebMethod.HttpPut
        Cloned.AddCookie "E", "cookie"
        Cloned.AddHeader "F", "header"
        Cloned.AddQuerystringParam "G", "querystring"
        Cloned.AddUrlSegment "H", "segment"
        
        .Expect(Request.Resource).ToEqual "resource/"
        .Expect(Request.Method).ToEqual WebMethod.HttpPost
        .Expect(WebHelpers.FindInKeyValues(Request.Cookies, "E")).ToBeUndefined
        .Expect(WebHelpers.FindInKeyValues(Request.Headers, "F")).ToBeUndefined
        .Expect(WebHelpers.FindInKeyValues(Request.QuerystringParams, "G")).ToBeUndefined
        .Expect(Request.UrlSegments.Exists("H")).ToEqual False
    End With
    
    ' Prepare
    ' @internal
    ' --------------------------------------------- '
    With Specs.It("Prepare should add ContentType, Accept, and ContentLength headers")
        Set Request = New WebRequest
        
        Request.Method = WebMethod.HttpPost
        Request.ContentType = "text/plain"
        Request.Accept = "text/csv"
        Request.ContentLength = 100
        
        .Expect(Request.Headers.Count).ToEqual 0
         
        Request.Prepare
        
        .Expect(Request.Headers.Count).ToBeGTE 3
        .Expect(WebHelpers.FindInKeyValues(Request.Headers, "Content-Type")).ToEqual "text/plain"
        .Expect(WebHelpers.FindInKeyValues(Request.Headers, "Accept")).ToEqual "text/csv"
        .Expect(WebHelpers.FindInKeyValues(Request.Headers, "Content-Length")).ToEqual "100"
    End With
    
    With Specs.It("Prepare should only add headers once")
        Set Request = New WebRequest
        
        Request.ContentType = "text/plain"
        Request.Accept = "text/csv"
        Request.ContentLength = 100
         
        Request.Prepare
        
        NumHeaders = Request.Headers.Count
        
        Request.Prepare
        
        .Expect(Request.Headers.Count).ToEqual NumHeaders
    End With
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    On Error Resume Next
    
    With Specs.It("AddBodyParameter should throw error if existing body isn't Dictionary")
        Set Request = New WebRequest
        
        Request.Body = "Howdy"
        Request.AddBodyParameter "Message", "Goodby"
        
        .Expect(Err.Number).ToEqual 11020 + vbObjectError
    End With
End Function


