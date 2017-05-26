Attribute VB_Name = "Specs_WebHelpers"
''
' Specs_WebHelpers
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for WebHelpers
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebHelpers"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs
    
    ' Contents:
    ' 1. Logging
    ' 2. Converters and encoding
    ' 3. Url handling
    ' 4. Object/Dictionary/Collection helpers
    ' 5. Request preparation / handling
    ' 6. Timing
    ' 7. Mac
    ' 8. Cryptography
    ' 9. Date/Time conversion
    ' Errors
    ' --------------------------------------------- '
    
    Dim JsonString As String
    Dim XmlString As String
    Dim Parsed As Object
    Dim Obj As Object
    Dim LocalDate As Date
    Dim Coll As Collection
    Dim Bytes() As Byte
    Dim Str As String
    Dim Encoded As String
    Dim Parts As Dictionary
    Dim Arr As Variant
    Dim Var As Variant
    Dim Strings() As String
    Dim OriginalDict As Dictionary
    Dim ClonedDict As Dictionary
    Dim OriginalColl As Collection
    Dim ClonedColl As Collection
    Dim KeyValue As Dictionary
    Dim KeyValues As Collection
    
    ' ============================================= '
    ' 1. Logging
    ' ============================================= '
    
    ' LogDebug
    ' LogError
    ' LogRequest
    ' LogResponse
    
    ' Obfuscate
    ' --------------------------------------------- '
    With Specs.It("should obfuscate string (with character option)")
        .Expect(WebHelpers.Obfuscate("secret")).ToEqual "******"
        .Expect(WebHelpers.Obfuscate("abc", "_")).ToEqual "___"
    End With

    ' ============================================= '
    ' 2. Converters and encoding
    ' ============================================= '
    
    ' ParseJson
    ' --------------------------------------------- '
    With Specs.It("should parse JSON")
        JsonString = "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2]}"
        Set Parsed = WebHelpers.ParseJson(JsonString)
        
        .Expect(Parsed).ToNotBeUndefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed("a")).ToEqual 1
            .Expect(Parsed("b")).ToEqual 3.14
            .Expect(Parsed("c")).ToEqual "Howdy!"
            .Expect(Parsed("d")).ToEqual True
            .Expect(Parsed("e").Count).ToEqual 2
        End If
        
        JsonString = "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""}]"
        Set Parsed = WebHelpers.ParseJson(JsonString)
        
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
    
    ' ConvertToJson
    ' --------------------------------------------- '
    With Specs.It("should convert to JSON")
        Set Obj = New Dictionary
        Obj.Add "a", 1
        Obj.Add "b", 3.14
        Obj.Add "c", "Howdy!"
        Obj.Add "d", True
        Obj.Add "e", Array(1, 2)
        Obj.Add "f", Empty
        Obj.Add "g", Null
        
        JsonString = WebHelpers.ConvertToJson(Obj)
        .Expect(JsonString).ToEqual "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2],""g"":null}"
        
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
        
        JsonString = WebHelpers.ConvertToJson(Coll)
        .Expect(JsonString).ToEqual "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""},null,null]"
    End With
    
    ' ParseUrlEncoded
    ' --------------------------------------------- '
    With Specs.It("should parse URL-encoded")
        Set Parsed = WebHelpers.ParseUrlEncoded("a=1&b=3.14&c=Howdy%21&d+%26+e=A+%2B+B")
        
        .Expect(Parsed("a")).ToEqual "1"
        .Expect(Parsed("b")).ToEqual "3.14"
        .Expect(Parsed("c")).ToEqual "Howdy!"
        .Expect(Parsed("d & e")).ToEqual "A + B"
    End With
    
    ' ConvertToUrlEncoded
    ' --------------------------------------------- '
    With Specs.It("should convert to URL-encoded")
        Set Obj = New Dictionary
        
        Obj.Add "a", 1
        Obj.Add "b", "Howdy!"
        Obj.Add "c & d", "A + B"
        
        Encoded = WebHelpers.ConvertToUrlEncoded(Obj)
        .Expect(Encoded).ToEqual "a=1&b=Howdy%21&c+%26+d=A+%2B+B"
    End With
    
    With Specs.It("should use region invariant numbers and dates")
        Set Obj = New Dictionary
        LocalDate = 38113.7973263889
        
        Obj.Add "a", 1000.123
        Obj.Add "b", LocalDate
        
        Encoded = WebHelpers.ConvertToUrlEncoded(Obj)
        
        ' Don't test hour/minute due to UTC offset
        .Expect(Encoded).ToMatch "a=1000.123&b=2004-05-06T"
        .Expect(Encoded).ToMatch "%3A09.000Z"
    End With
   
    ' ParseXml
    ' --------------------------------------------- '
'    With Specs.It("[Windows-only] should parse XML")
'        Set Parsed = WebHelpers.ParseXml("<Point><X>1.23</X><Y>4.56</Y></Point>")
'
'        .Expect(Parsed.FirstChild.SelectSingleNode("X").Text).ToEqual "1.23"
'        .Expect(Parsed.FirstChild.SelectSingleNode("Y").Text).ToEqual "4.56"
'    End With
    
    ' ConvertToXml
    ' --------------------------------------------- '
'    With Specs.It("[Windows-only] should convert to XML")
'        XmlString = "<Point><X>1.23</X><Y>4.56</Y></Point>"
'        Set Obj = CreateObject("MSXML2.DOMDocument")
'        Obj.Async = False
'        Obj.LoadXML XmlString
'
'        Encoded = WebHelpers.ConvertToXml(Obj)
'        .Expect(Encoded).ToEqual XmlString
'    End With
    
    ' ParseByFormat
    ' --------------------------------------------- '
    
    ' ConvertToFormat
    ' --------------------------------------------- '
    
    ' UrlEncode
    ' --------------------------------------------- '
    With Specs.It("should url-encode (Default = StrictUrlEncoding)")
        .Expect(WebHelpers.UrlEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZ")).ToEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        .Expect(WebHelpers.UrlEncode("abcdefghijklmnopqrstuvwxyz")).ToEqual "abcdefghijklmnopqrstuvwxyz"
        .Expect(WebHelpers.UrlEncode("1234567890")).ToEqual "1234567890"
        
        .Expect(WebHelpers.UrlEncode("-")).ToEqual "-"
        .Expect(WebHelpers.UrlEncode(".")).ToEqual "."
        .Expect(WebHelpers.UrlEncode("_")).ToEqual "_"
        .Expect(WebHelpers.UrlEncode("~")).ToEqual "~"
        
        .Expect(WebHelpers.UrlEncode("%")).ToEqual "%25"
        .Expect(WebHelpers.UrlEncode(" ")).ToEqual "%20"

        .Expect(WebHelpers.UrlEncode("!")).ToEqual "%21"
        .Expect(WebHelpers.UrlEncode("#")).ToEqual "%23"
        .Expect(WebHelpers.UrlEncode("$")).ToEqual "%24"
        .Expect(WebHelpers.UrlEncode("&")).ToEqual "%26"
        .Expect(WebHelpers.UrlEncode("'")).ToEqual "%27"
        .Expect(WebHelpers.UrlEncode("(")).ToEqual "%28"
        .Expect(WebHelpers.UrlEncode(")")).ToEqual "%29"
        .Expect(WebHelpers.UrlEncode("*")).ToEqual "%2A"
        .Expect(WebHelpers.UrlEncode("+")).ToEqual "%2B"
        .Expect(WebHelpers.UrlEncode(",")).ToEqual "%2C"
        .Expect(WebHelpers.UrlEncode("/")).ToEqual "%2F"
        .Expect(WebHelpers.UrlEncode(":")).ToEqual "%3A"
        .Expect(WebHelpers.UrlEncode(";")).ToEqual "%3B"
        .Expect(WebHelpers.UrlEncode("=")).ToEqual "%3D"
        .Expect(WebHelpers.UrlEncode("?")).ToEqual "%3F"
        .Expect(WebHelpers.UrlEncode("@")).ToEqual "%40"
        .Expect(WebHelpers.UrlEncode("[")).ToEqual "%5B"
        .Expect(WebHelpers.UrlEncode("]")).ToEqual "%5D"
        .Expect(WebHelpers.UrlEncode("^")).ToEqual "%5E"
        .Expect(WebHelpers.UrlEncode("`")).ToEqual "%60"
        .Expect(WebHelpers.UrlEncode("|")).ToEqual "%7C"
    End With
    
    With Specs.It("should url-encode (FormUrlEncoding)")
        .Expect(WebHelpers.UrlEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZ", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        .Expect(WebHelpers.UrlEncode("abcdefghijklmnopqrstuvwxyz", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "abcdefghijklmnopqrstuvwxyz"
        .Expect(WebHelpers.UrlEncode("1234567890", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "1234567890"
        
        .Expect(WebHelpers.UrlEncode("-", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "-"
        .Expect(WebHelpers.UrlEncode(".", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "."
        .Expect(WebHelpers.UrlEncode("_", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "_"
        .Expect(WebHelpers.UrlEncode("~", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%7E"
        
        .Expect(WebHelpers.UrlEncode("%", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%25"
        .Expect(WebHelpers.UrlEncode(" ", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "+"

        .Expect(WebHelpers.UrlEncode("!", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%21"
        .Expect(WebHelpers.UrlEncode("#", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%23"
        .Expect(WebHelpers.UrlEncode("$", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%24"
        .Expect(WebHelpers.UrlEncode("&", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%26"
        .Expect(WebHelpers.UrlEncode("'", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%27"
        .Expect(WebHelpers.UrlEncode("(", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%28"
        .Expect(WebHelpers.UrlEncode(")", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%29"
        .Expect(WebHelpers.UrlEncode("*", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "*"
        .Expect(WebHelpers.UrlEncode("+", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%2B"
        .Expect(WebHelpers.UrlEncode(",", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%2C"
        .Expect(WebHelpers.UrlEncode("/", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%2F"
        .Expect(WebHelpers.UrlEncode(":", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%3A"
        .Expect(WebHelpers.UrlEncode(";", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%3B"
        .Expect(WebHelpers.UrlEncode("=", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%3D"
        .Expect(WebHelpers.UrlEncode("?", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%3F"
        .Expect(WebHelpers.UrlEncode("@", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%40"
        .Expect(WebHelpers.UrlEncode("[", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%5B"
        .Expect(WebHelpers.UrlEncode("]", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%5D"
        .Expect(WebHelpers.UrlEncode("^", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%5E"
        .Expect(WebHelpers.UrlEncode("`", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%60"
        .Expect(WebHelpers.UrlEncode("|", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "%7C"
    End With

    With Specs.It("should url-encode (QueryUrlEncoding)")
        .Expect(WebHelpers.UrlEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZ", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        .Expect(WebHelpers.UrlEncode("abcdefghijklmnopqrstuvwxyz", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "abcdefghijklmnopqrstuvwxyz"
        .Expect(WebHelpers.UrlEncode("1234567890", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "1234567890"
        
        .Expect(WebHelpers.UrlEncode("-", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "-"
        .Expect(WebHelpers.UrlEncode(".", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "."
        .Expect(WebHelpers.UrlEncode("_", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "_"
        .Expect(WebHelpers.UrlEncode("~", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%7E"
        
        .Expect(WebHelpers.UrlEncode("%", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%25"
        .Expect(WebHelpers.UrlEncode(" ", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%20"

        .Expect(WebHelpers.UrlEncode("!", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%21"
        .Expect(WebHelpers.UrlEncode("#", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%23"
        .Expect(WebHelpers.UrlEncode("$", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%24"
        .Expect(WebHelpers.UrlEncode("&", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%26"
        .Expect(WebHelpers.UrlEncode("'", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%27"
        .Expect(WebHelpers.UrlEncode("(", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%28"
        .Expect(WebHelpers.UrlEncode(")", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%29"
        .Expect(WebHelpers.UrlEncode("*", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%2A"
        .Expect(WebHelpers.UrlEncode("+", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%2B"
        .Expect(WebHelpers.UrlEncode(",", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%2C"
        .Expect(WebHelpers.UrlEncode("/", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%2F"
        .Expect(WebHelpers.UrlEncode(":", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%3A"
        .Expect(WebHelpers.UrlEncode(";", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%3B"
        .Expect(WebHelpers.UrlEncode("=", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%3D"
        .Expect(WebHelpers.UrlEncode("?", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%3F"
        .Expect(WebHelpers.UrlEncode("@", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%40"
        .Expect(WebHelpers.UrlEncode("[", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%5B"
        .Expect(WebHelpers.UrlEncode("]", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%5D"
        .Expect(WebHelpers.UrlEncode("^", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%5E"
        .Expect(WebHelpers.UrlEncode("`", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%60"
        .Expect(WebHelpers.UrlEncode("|", EncodingMode:=UrlEncodingMode.QueryUrlEncoding)).ToEqual "%7C"
    End With

    With Specs.It("should url-encode (CookieUrlEncoding)")
        .Expect(WebHelpers.UrlEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZ", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        .Expect(WebHelpers.UrlEncode("abcdefghijklmnopqrstuvwxyz", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "abcdefghijklmnopqrstuvwxyz"
        .Expect(WebHelpers.UrlEncode("1234567890", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "1234567890"
        
        .Expect(WebHelpers.UrlEncode("-", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "-"
        .Expect(WebHelpers.UrlEncode(".", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "."
        .Expect(WebHelpers.UrlEncode("_", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "_"
        .Expect(WebHelpers.UrlEncode("~", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "~"
        
        ' Note: "%" is allowed in spec, but is currently excluded due to parsing issues
        .Expect(WebHelpers.UrlEncode("%", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "%25"
        
        .Expect(WebHelpers.UrlEncode("""", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "%22"
        .Expect(WebHelpers.UrlEncode(" ", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "%20"
        
        .Expect(WebHelpers.UrlEncode("!", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "!"
        .Expect(WebHelpers.UrlEncode("#", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "#"
        .Expect(WebHelpers.UrlEncode("$", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "$"
        .Expect(WebHelpers.UrlEncode("&", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "&"
        .Expect(WebHelpers.UrlEncode("'", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "'"
        .Expect(WebHelpers.UrlEncode("(", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "("
        .Expect(WebHelpers.UrlEncode(")", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual ")"
        .Expect(WebHelpers.UrlEncode("*", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "*"
        .Expect(WebHelpers.UrlEncode("+", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "+"
        .Expect(WebHelpers.UrlEncode(",", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "%2C"
        .Expect(WebHelpers.UrlEncode("/", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "/"
        .Expect(WebHelpers.UrlEncode(":", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual ":"
        .Expect(WebHelpers.UrlEncode(";", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "%3B"
        .Expect(WebHelpers.UrlEncode("<", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "<"
        .Expect(WebHelpers.UrlEncode("=", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "="
        .Expect(WebHelpers.UrlEncode(">", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual ">"
        .Expect(WebHelpers.UrlEncode("?", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "?"
        .Expect(WebHelpers.UrlEncode("@", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "@"
        .Expect(WebHelpers.UrlEncode("[", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "["
        .Expect(WebHelpers.UrlEncode("]", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "]"
        .Expect(WebHelpers.UrlEncode("^", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "^"
        .Expect(WebHelpers.UrlEncode("`", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "`"
        .Expect(WebHelpers.UrlEncode("{", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "{"
        .Expect(WebHelpers.UrlEncode("|", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "|"
        .Expect(WebHelpers.UrlEncode("}", EncodingMode:=UrlEncodingMode.CookieUrlEncoding)).ToEqual "}"
    End With

    With Specs.It("should url-encode (PathUrlEncoding)")
        .Expect(WebHelpers.UrlEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZ", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        .Expect(WebHelpers.UrlEncode("abcdefghijklmnopqrstuvwxyz", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "abcdefghijklmnopqrstuvwxyz"
        .Expect(WebHelpers.UrlEncode("1234567890", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "1234567890"
        
        .Expect(WebHelpers.UrlEncode("-", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "-"
        .Expect(WebHelpers.UrlEncode(".", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "."
        .Expect(WebHelpers.UrlEncode("_", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "_"
        .Expect(WebHelpers.UrlEncode("~", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "~"
        
        .Expect(WebHelpers.UrlEncode("%", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%25"
        .Expect(WebHelpers.UrlEncode(" ", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%20"

        .Expect(WebHelpers.UrlEncode("!", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "!"
        .Expect(WebHelpers.UrlEncode("#", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%23"
        .Expect(WebHelpers.UrlEncode("$", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "$"
        .Expect(WebHelpers.UrlEncode("&", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "&"
        .Expect(WebHelpers.UrlEncode("'", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "'"
        .Expect(WebHelpers.UrlEncode("(", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "("
        .Expect(WebHelpers.UrlEncode(")", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual ")"
        .Expect(WebHelpers.UrlEncode("*", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "*"
        .Expect(WebHelpers.UrlEncode("+", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "+"
        .Expect(WebHelpers.UrlEncode(",", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual ","
        .Expect(WebHelpers.UrlEncode("/", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%2F"
        .Expect(WebHelpers.UrlEncode(":", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual ":"
        .Expect(WebHelpers.UrlEncode(";", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual ";"
        .Expect(WebHelpers.UrlEncode("=", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "="
        .Expect(WebHelpers.UrlEncode("?", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%3F"
        .Expect(WebHelpers.UrlEncode("@", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "@"
        .Expect(WebHelpers.UrlEncode("[", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%5B"
        .Expect(WebHelpers.UrlEncode("]", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%5D"
        .Expect(WebHelpers.UrlEncode("^", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%5E"
        .Expect(WebHelpers.UrlEncode("`", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%60"
        .Expect(WebHelpers.UrlEncode("|", EncodingMode:=UrlEncodingMode.PathUrlEncoding)).ToEqual "%7C"
    End With
    
    With Specs.It("DEPRECATED should url-encode with SpaceAsPlus")
        .Expect(WebHelpers.UrlEncode("A + B", SpaceAsPlus:=True)).ToEqual "A+%2B+B"
    End With
    
    ' UrlDecode
    ' --------------------------------------------- '
    With Specs.It("should url-decode string (Default = StrictUrlEncoding)")
        .Expect(WebHelpers.UrlDecode("ABCDEFGHIJKLMNOPQRSTUVWXYZ")).ToEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        .Expect(WebHelpers.UrlDecode("abcdefghijklmnopqrstuvwxyz")).ToEqual "abcdefghijklmnopqrstuvwxyz"
        .Expect(WebHelpers.UrlDecode("1234567890")).ToEqual "1234567890"
        
        .Expect(WebHelpers.UrlDecode("%0A")).ToEqual vbLf
        .Expect(WebHelpers.UrlDecode("%25")).ToEqual "%"
        .Expect(WebHelpers.UrlDecode("%7E")).ToEqual "~"
    End With
    
    With Specs.It("should url-decode string (FormUrlEncoding)")
        .Expect(WebHelpers.UrlDecode("1%20+%202", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "1   2"
    End With
    
    With Specs.It("should url-decode string (QueryUrlEncoding)")
        .Expect(WebHelpers.UrlDecode("1%20+%202", EncodingMode:=UrlEncodingMode.FormUrlEncoding)).ToEqual "1   2"
    End With
    
    With Specs.It("DEPRECATED should url-decode with PlusAsSpace as default")
        .Expect(WebHelpers.UrlDecode("1%20+%202")).ToEqual "1   2"
    End With
    
    ' Base64Encode
    ' --------------------------------------------- '
    With Specs.It("should Base64 encode string")
        .Expect(WebHelpers.Base64Encode("Howdy!")).ToEqual "SG93ZHkh"
    End With
    
    ' Base64Decode
    ' --------------------------------------------- '
    With Specs.It("should Base64 decode string (with or without padding)")
        .Expect(WebHelpers.Base64Decode("SG93ZHkh")).ToEqual "Howdy!"
        
        ' The following implicitly has padding of "=" and "==" at end, base-64 decoding should handle this
        .Expect(WebHelpers.Base64Decode("SG93ZHk")).ToEqual "Howdy"
        .Expect(WebHelpers.Base64Decode("eyJzdWIiOjEyMzQ1Njc4OTAsIm5hbWUiOiJKb2huIERvZSIsImFkbWluIjp0cnVlfQ")).ToEqual _
            "{""sub"":1234567890,""name"":""John Doe"",""admin"":true}"
    End With
    
    ' RegisterConverter
    ' --------------------------------------------- '
    With Specs.It("RegisterConverter should register and use converter")
        WebHelpers.RegisterConverter "custom-a", "X-a", "Specs_WebHelpers.SimpleConverter", "Specs_WebHelpers.SimpleParser"
        
        Set Obj = New Dictionary
        Obj.Add "message", "Howdy!"
        
        JsonString = WebHelpers.ConvertToFormat(Obj, WebFormat.Custom, "custom-a")
        .Expect(JsonString).ToEqual "{""message"":""Howdy!"",""response"":""Goodbye!""}"
    End With
    
    With Specs.It("RegisterConverter should register and use converter with instance")
        Dim Converter As New SpecConverter
        WebHelpers.RegisterConverter "custom-b", "X-b", "ConvertToCustom", "ParseCustom", Converter
        
        Set Parsed = WebHelpers.ParseByFormat("{""message"":""Howdy!""}", WebFormat.Custom, "custom-b")
        .Expect(Parsed).ToNotBeUndefined
        .Expect(Parsed("response")).ToEqual "Goodbye!"
    End With
    
    With Specs.It("RegisterConverter should register and use converter with Binary ParseType")
        WebHelpers.RegisterConverter "custom-c", "X-c", "Specs_WebHelpers.SimpleConverter", "Specs_WebHelpers.ComplexParser", ParseType:="Binary"
        
        Str = "Howdy!"
        Bytes = Str
        
        Set Parsed = WebHelpers.ParseByFormat("", WebFormat.Custom, "custom-c", Array(72, 111, 119, 100, 121, 33))
        .Expect(Parsed).ToNotBeUndefined
        .Expect(Parsed("message")).ToEqual "Howdy!"
        .Expect(Parsed("response")).ToEqual "Goodbye!"
    End With
    
    ' ============================================= '
    ' 3. Url handling
    ' ============================================= '
    
    ' JoinUrl
    ' --------------------------------------------- '
    With Specs.It("JoinUrl should join url with /")
        .Expect(WebHelpers.JoinUrl("a", "b")).ToEqual "a/b"
        .Expect(WebHelpers.JoinUrl("a/", "b")).ToEqual "a/b"
        .Expect(WebHelpers.JoinUrl("a", "/b")).ToEqual "a/b"
        .Expect(WebHelpers.JoinUrl("a/", "/b")).ToEqual "a/b"
    End With
    
    With Specs.It("JoinUrl should not join blank urls with /")
        .Expect(WebHelpers.JoinUrl("", "b")).ToEqual "b"
        .Expect(WebHelpers.JoinUrl("a", "")).ToEqual "a"
    End With
    
    ' GetUrlParts
    ' --------------------------------------------- '
    With Specs.It("should extract parts from url")
        Set Parts = WebHelpers.GetUrlParts("https://www.google.com/dir/1/2/search.html?message=Howdy%20World!&other=123#hash")
        
        .Expect(Parts("Protocol")).ToEqual "https"
        .Expect(Parts("Host")).ToEqual "www.google.com"
        .Expect(Parts("Port")).ToEqual "443"
        .Expect(Parts("Path")).ToEqual "/dir/1/2/search.html"
        .Expect(Parts("Querystring")).ToEqual "message=Howdy%20World!&other=123"
        .Expect(Parts("Hash")).ToEqual "hash"
        
        Set Parts = WebHelpers.GetUrlParts("localhost:3000/dir/1/2/page%202.html?message=Howdy%20World!&other=123#hash")
        
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
    
    ' CloneDictionary
    ' --------------------------------------------- '
    With Specs.It("should clone Dictionary")
        Set OriginalDict = New Dictionary
        OriginalDict.Add "a", "abc"
        OriginalDict.Add "b", 123
        OriginalDict.Add "c", 3.14
        OriginalDict.Add "d", True
        OriginalDict.Add "e", Array(3, 2, 1)
        OriginalDict.Add "f", New Dictionary
        OriginalDict("f").Add "message", "Howdy!"
        
        Set ClonedDict = WebHelpers.CloneDictionary(OriginalDict)
        
        .Expect(ClonedDict("a")).ToEqual "abc"
        .Expect(ClonedDict("b")).ToEqual 123
        .Expect(ClonedDict("c")).ToEqual 3.14
        .Expect(ClonedDict("d")).ToEqual True
        .Expect(ClonedDict("e")(0)).ToEqual 3
        .Expect(ClonedDict("f")("message")).ToEqual "Howdy!"
        
        ClonedDict("b") = 456
        ClonedDict.Remove "d"
        .Expect(OriginalDict("b")).ToEqual 123
        .Expect(OriginalDict.Exists("d")).ToEqual True
    End With
    
    ' CloneCollection
    ' --------------------------------------------- '
    With Specs.It("should clone Collection")
        Set OriginalColl = New Collection
        OriginalColl.Add "abc"
        OriginalColl.Add 123
        OriginalColl.Add 3.14
        OriginalColl.Add True
        OriginalColl.Add Array(3, 2, 1)
        OriginalColl.Add New Dictionary
        OriginalColl(6).Add "message", "Howdy!"
        
        Set ClonedColl = WebHelpers.CloneCollection(OriginalColl)
        
        .Expect(ClonedColl(1)).ToEqual "abc"
        .Expect(ClonedColl(2)).ToEqual 123
        .Expect(ClonedColl(3)).ToEqual 3.14
        .Expect(ClonedColl(4)).ToEqual True
        .Expect(ClonedColl(5)(0)).ToEqual 3
        .Expect(ClonedColl(6)("message")).ToEqual "Howdy!"
        
        ClonedColl.Remove 4
        .Expect(OriginalColl.Count).ToEqual 6
    End With
    
    ' CreateKeyValue
    ' --------------------------------------------- '
    With Specs.It("should create Key-Value Dictionary")
        Set KeyValue = WebHelpers.CreateKeyValue("abc", 123)
        .Expect(KeyValue("Key")).ToEqual "abc"
        .Expect(KeyValue("Value")).ToEqual 123
    End With
    
    ' FindInKeyValues
    ' --------------------------------------------- '
    With Specs.It("should find Value by Key in Key-Values")
        Set KeyValues = New Collection
        KeyValues.Add WebHelpers.CreateKeyValue("a", 123)
        KeyValues.Add WebHelpers.CreateKeyValue("b", 456)
        KeyValues.Add WebHelpers.CreateKeyValue("c", 789)
        
        .Expect(WebHelpers.FindInKeyValues(KeyValues, "b")).ToEqual 456
        .Expect(WebHelpers.FindInKeyValues(KeyValues, "d")).ToBeEmpty
    End With
    
    ' AddOrReplaceInKeyValues
    ' --------------------------------------------- '
    With Specs.It("should add or replace (with retained order) in Key-Values")
        Set KeyValues = New Collection
        KeyValues.Add WebHelpers.CreateKeyValue("a", 123)
        KeyValues.Add WebHelpers.CreateKeyValue("b", 456)
        KeyValues.Add WebHelpers.CreateKeyValue("c", 789)
        
        WebHelpers.AddOrReplaceInKeyValues KeyValues, "a", "abc"
        WebHelpers.AddOrReplaceInKeyValues KeyValues, "b", "def"
        WebHelpers.AddOrReplaceInKeyValues KeyValues, "c", "ghi"
        WebHelpers.AddOrReplaceInKeyValues KeyValues, "d", "jkl"
        
        .Expect(KeyValues.Count).ToEqual 4
        .Expect(KeyValues(1)("Key")).ToEqual "a"
        .Expect(KeyValues(2)("Key")).ToEqual "b"
        .Expect(KeyValues(3)("Key")).ToEqual "c"
        .Expect(KeyValues(4)("Key")).ToEqual "d"
        
        .Expect(KeyValues(2)("Value")).ToEqual "def"
        .Expect(KeyValues(4)("Value")).ToEqual "jkl"
    End With
    
    ' ============================================= '
    ' 5. Request preparation / handling
    ' ============================================= '
    
    ' FormatToMediaType
    With Specs.It("FormatToMediaType should handle custom converters")
        .Expect(WebHelpers.FormatToMediaType(WebFormat.Custom, "custom-a")).ToEqual "X-a"
        .Expect(WebHelpers.FormatToMediaType(WebFormat.Custom, "custom-b")).ToEqual "X-b"
        .Expect(WebHelpers.FormatToMediaType(WebFormat.Custom, "custom-c")).ToEqual "X-c"
    End With
    
    ' MethodToName
    ' AddAsyncRequest
    ' GetAsyncRequest
    ' RemoveAsyncRequest
    
    ' ============================================= '
    ' 6. Timing
    ' ============================================= '
    
    ' StartTimeoutTimer
    ' StopTimeoutTimer
    ' OnTimeoutTimerExpired
    
    ' ============================================= '
    ' 7. Mac
    ' ============================================= '
    
    ' ExecuteInShell
    
    ' PrepareTextForShell
    ' --------------------------------------------- '
    With Specs.It("should prepare text for shell")
        .Expect(WebHelpers.PrepareTextForShell("""message""")).ToEqual """\""message\"""""
        .Expect(WebHelpers.PrepareTextForShell("!abc!123!")).ToEqual "'!'""abc""'!'""123""'!'"
        .Expect(WebHelpers.PrepareTextForShell("`!$\""%")).ToEqual """\`""'!'""\$\\\""\%"""
    End With
    
    ' ============================================= '
    ' 8. Cryptography
    ' ============================================= '
    
    ' HMACSHA1
    ' --------------------------------------------- '
    With Specs.It("should calculate HMAC with SHA1 algorithm")
        .Expect(WebHelpers.HMACSHA1("test", "secret")).ToEqual "1aa349585ed7ecbd3b9c486a30067e395ca4b356"
        .Expect(WebHelpers.HMACSHA1("123456789", "987654321")).ToEqual "eea1a8e956b1b26067e6d0bef57e54490b8892a9"

        .Expect(WebHelpers.HMACSHA1("test", "secret", "Base64")).ToEqual "GqNJWF7X7L07nEhqMAZ+OVyks1Y="
        .Expect(WebHelpers.HMACSHA1("123456789", "987654321", "Base64")).ToEqual "7qGo6VaxsmBn5tC+9X5USQuIkqk="
    End With
    
    ' HMACSHA256
    ' --------------------------------------------- '
    With Specs.It("should calculate HMAC with SHA256 algorithm")
        .Expect(WebHelpers.HMACSHA256("test", "secret")).ToEqual "0329a06b62cd16b33eb6792be8c60b158d89a2ee3a876fce9a881ebb488c0914"
        .Expect(WebHelpers.HMACSHA256("123456789", "987654321")).ToEqual "3122584687113ac66d3c2f3c3518c789eef536a298121e0dbc82fc8fe7621e73"
        
        .Expect(WebHelpers.HMACSHA256("test", "secret", "Base64")).ToEqual "Aymga2LNFrM+tnkr6MYLFY2Jou46h2/Omogeu0iMCRQ="
        .Expect(WebHelpers.HMACSHA256("123456789", "987654321", "Base64")).ToEqual "MSJYRocROsZtPC88NRjHie71NqKYEh4NvIL8j+diHnM="
    End With
    
    ' MD5
    ' --------------------------------------------- '
    With Specs.It("should MD5 hash string")
        .Expect(WebHelpers.MD5("test")).ToEqual "098f6bcd4621d373cade4e832627b4f6"
        .Expect(WebHelpers.MD5("123456789")).ToEqual "25f9e794323b453885f5181f1b624d0b"
        .Expect(WebHelpers.MD5("Mufasa:testrealm@host.com:Circle Of Life")).ToEqual "939e7578ed9e3c518a452acee763bce9"
        
        .Expect(WebHelpers.MD5("test", "Base64")).ToEqual "CY9rzUYh03PK3k6DJie09g=="
        .Expect(WebHelpers.MD5("123456789", "Base64")).ToEqual "JfnnlDI7RTiF9RgfG2JNCw=="
        .Expect(WebHelpers.MD5("Mufasa:testrealm@host.com:Circle Of Life", "Base64")).ToEqual "k551eO2ePFGKRSrO52O86Q=="
    End With
    
    ' CreateNonce
    ' --------------------------------------------- '
    With Specs.It("should create Nonce of specified length")
        .Expect(Len(WebHelpers.CreateNonce)).ToEqual 32
        .Expect(Len(WebHelpers.CreateNonce(20))).ToEqual 20
    End With
    
    
    ' ============================================= '
    ' 9. Date/Time conversion
    ' ============================================= '
    
    ' ISO to UTC (all ISO dates in daylight saving time)
    ' --------------------------------------------- '
    With Specs.It("should handle offset in ISO date")
        .Expect(WebHelpers.ConvertToUtc(WebHelpers.ParseIso("2017-05-01T02:00:00.000+02:00"))).ToEqual DateValue("2017-05-01") + TimeValue("00:00:00") ' 02:00:00 in Berlin => 00:00 (same day) UTC
        .Expect(WebHelpers.ConvertToUtc(WebHelpers.ParseIso("2017-05-01T00:00:00.000+02:00"))).ToEqual DateValue("2017-04-30") + TimeValue("22:00:00") ' 00:00:00 in Berlin => 22:00 (prev. day) UTC
        .Expect(WebHelpers.ConvertToUtc(WebHelpers.ParseIso("2017-04-30T20:00:00.000-04:00"))).ToEqual DateValue("2017-05-01") + TimeValue("00:00:00") ' 20:00:00 in New York => 00:00 (next day) UTC
    End With
    
    ' ISO to UTC
    ' --------------------------------------------- '
    With Specs.It("should convert ISO dates in UTC to UTC")
        .Expect(WebHelpers.ConvertToUtc(WebHelpers.ParseIso("2017-05-01T00:00:00.000Z"))).ToEqual DateValue("2017-05-01 00:00:00") + TimeValue("00:00:00")
        .Expect(WebHelpers.ConvertToUtc(WebHelpers.ParseIso("2017-05-01T00:00:00.000+00:00"))).ToEqual DateValue("2017-05-01 00:00:00") + TimeValue("00:00:00")
    End With
    
    
    
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    On Error Resume Next
End Function

Function SimpleConverter(Body As Variant) As String
    Body.Add "response", "Goodbye!"
    SimpleConverter = WebHelpers.ConvertToJson(Body)
End Function
Function SimpleParser(Content As String) As Object
    Set SimpleParser = WebHelpers.ParseJson(Content)
    SimpleParser.Add "response", "Goodbye!"
End Function
Function ComplexParser(Body As Variant) As Object
    Dim Content As String
    Dim i As Integer
    
    For i = LBound(Body) To UBound(Body)
        Content = Content & Chr(Body(i))
    Next i
    
    Set ComplexParser = WebHelpers.ParseJson("{""message"":""" & Content & """}")
    ComplexParser.Add "response", "Goodbye!"
End Function
