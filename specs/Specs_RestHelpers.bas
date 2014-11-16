Attribute VB_Name = "Specs_RestHelpers"
''
' Specs_RestHelpers
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
    ' 7. Mac
    ' 8. Cryptography
    ' --------------------------------------------- '
    
    Dim JSONString As String
    Dim XMLString As String
    Dim Parsed As Object
    Dim Obj As Object
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
        .Expect(RestHelpers.Obfuscate("secret")).ToEqual "******"
        .Expect(RestHelpers.Obfuscate("abc", "_")).ToEqual "___"
    End With

    ' ============================================= '
    ' 2. Converters and encoding
    ' ============================================= '
    
    ' ParseJson
    ' --------------------------------------------- '
    With Specs.It("should parse JSON")
        JSONString = "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2]}"
        Set Parsed = RestHelpers.ParseJSON(JSONString)
        
        .Expect(Parsed).ToNotBeUndefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed("a")).ToEqual 1
            .Expect(Parsed("b")).ToEqual 3.14
            .Expect(Parsed("c")).ToEqual "Howdy!"
            .Expect(Parsed("d")).ToEqual True
            .Expect(Parsed("e").Count).ToEqual 2
        End If
        
        JSONString = "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""}]"
        Set Parsed = RestHelpers.ParseJSON(JSONString)
        
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
        
        JSONString = RestHelpers.ConvertToJSON(Obj)
        .Expect(JSONString).ToEqual "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2],""f"":null,""g"":null}"
        
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
        
        JSONString = RestHelpers.ConvertToJSON(Coll)
        .Expect(JSONString).ToEqual "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""},null,null]"
    End With
    
    ' ParseUrlEncoded
    ' --------------------------------------------- '
    With Specs.It("should parse URL-encoded")
        Set Parsed = RestHelpers.ParseUrlEncoded("a=1&b=3.14&c=Howdy%21&d+%26+e=A+%2B+B")
        
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
        Obj.Add "b", 3.14
        Obj.Add "c", "Howdy!"
        Obj.Add "d & e", "A + B"
        
        Encoded = RestHelpers.ConvertToUrlEncoded(Obj)
        .Expect(Encoded).ToEqual "a=1&b=3.14&c=Howdy!&d+%26+e=A+%2B+B"
    End With
   
#If Mac Then
#Else
    ' ParseXml
    ' --------------------------------------------- '
    With Specs.It("should parse XML")
        Set Parsed = RestHelpers.ParseXML("<Point><X>1.23</X><Y>4.56</Y></Point>")
        
        .Expect(Parsed.FirstChild.SelectSingleNode("X").Text).ToEqual "1.23"
        .Expect(Parsed.FirstChild.SelectSingleNode("Y").Text).ToEqual "4.56"
    End With
    
    ' ConvertToXml
    ' --------------------------------------------- '
    With Specs.It("should convert to XML")
        XMLString = "<Point><X>1.23</X><Y>4.56</Y></Point>"
        Set Obj = CreateObject("MSXML2.DOMDocument")
        Obj.Async = False
        Obj.LoadXML XMLString

        Encoded = RestHelpers.ConvertToXML(Obj)
        .Expect(Encoded).ToEqual XMLString
    End With
#End If
    
    ' ParseByFormat
    ' --------------------------------------------- '
    
    ' ConvertToFormat
    ' --------------------------------------------- '
    
    ' UrlEncode
    ' --------------------------------------------- '
    With Specs.It("should url-encode string (with space as plus and encode unsafe options)")
        .Expect(RestHelpers.UrlEncode("$&+,/:;=?@", EncodeUnsafe:=False)).ToEqual "%24%26%2B%2C%2F%3A%3B%3D%3F%40"
        .Expect(RestHelpers.UrlEncode(" ""<>#%{}|\^~[]`")).ToEqual "%20%22%3C%3E%23%25%7B%7D%7C%5C%5E%7E%5B%5D%60"
        .Expect(RestHelpers.UrlEncode("A + B")).ToEqual "A%20%2B%20B"
        .Expect(RestHelpers.UrlEncode("A + B", SpaceAsPlus:=True)).ToEqual "A+%2B+B"
    End With
    
    ' UrlDecode
    ' --------------------------------------------- '
    With Specs.It("should url-decode string")
        .Expect(RestHelpers.UrlDecode("+%20%21%22%23%24%25%26%27")).ToEqual "  !""#$%&'"
        .Expect(RestHelpers.UrlDecode("A%20%2B%20B")).ToEqual "A + B"
        .Expect(RestHelpers.UrlDecode("A+%2B+B")).ToEqual "A + B"
    End With
    
    ' Base64Encode
    ' --------------------------------------------- '
    With Specs.It("should Base64 encode string")
        .Expect(RestHelpers.Base64Encode("Howdy!")).ToEqual "SG93ZHkh"
    End With
    
    ' RegisterConverter
    ' --------------------------------------------- '
    With Specs.It("RegisterConverter should register and use converter")
        RestHelpers.RegisterConverter "custom-a", "X-a", "Specs_RestHelpers.SimpleConverter", "Specs_RestHelpers.SimpleParser"
        
        Set Obj = New Dictionary
        Obj.Add "message", "Howdy!"
        
        JSONString = RestHelpers.ConvertToFormat(Obj, WebFormat.custom, "custom-a")
        .Expect(JSONString).ToEqual "{""message"":""Howdy!"",""response"":""Goodbye!""}"
    End With
    
    With Specs.It("RegisterConverter should register and use converter with instance")
        Dim Converter As New SpecConverter
        RestHelpers.RegisterConverter "custom-b", "X-b", "ConvertToCustom", "ParseCustom", Converter
        
        Set Parsed = RestHelpers.ParseByFormat("{""message"":""Howdy!""}", WebFormat.custom, "custom-b")
        .Expect(Parsed).ToNotBeUndefined
        .Expect(Parsed("response")).ToEqual "Goodbye!"
    End With
    
    With Specs.It("RegisterConverter should register and use converter with Binary ParseType")
        RestHelpers.RegisterConverter "custom-c", "X-c", "Specs_RestHelpers.SimpleConverter", "Specs_RestHelpers.ComplexParser", ParseType:="Binary"
        
        Str = "Howdy!"
        Bytes = Str
        
        Set Parsed = RestHelpers.ParseByFormat("", WebFormat.custom, "custom-c", Array(72, 111, 119, 100, 121, 33))
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
        .Expect(RestHelpers.JoinUrl("a", "b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a/", "b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a", "/b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a/", "/b")).ToEqual "a/b"
    End With
    
    With Specs.It("JoinUrl should not join blank urls with /")
        .Expect(RestHelpers.JoinUrl("", "b")).ToEqual "b"
        .Expect(RestHelpers.JoinUrl("a", "")).ToEqual "a"
    End With
    
    ' UrlParts
    ' --------------------------------------------- '
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
    
    ' IsArray
    ' --------------------------------------------- '
    With Specs.It("should check for Variant and Typed arrays")
        Arr = Array(1, 2, 3)
        Erase Bytes
        
        .Expect(RestHelpers.IsArray(Arr)).ToEqual True
        .Expect(RestHelpers.IsArray(Bytes)).ToEqual True
        .Expect(RestHelpers.IsArray(Strings)).ToEqual True
        .Expect(RestHelpers.IsArray(Var)).ToEqual False
        .Expect(RestHelpers.IsArray(New Collection)).ToEqual False
    End With
    
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
        
        Set ClonedDict = RestHelpers.CloneDictionary(OriginalDict)
        
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
        
        Set ClonedColl = RestHelpers.CloneCollection(OriginalColl)
        
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
        Set KeyValue = RestHelpers.CreateKeyValue("abc", 123)
        .Expect(KeyValue("Key")).ToEqual "abc"
        .Expect(KeyValue("Value")).ToEqual 123
    End With
    
    ' FindInKeyValues
    ' --------------------------------------------- '
    With Specs.It("should find Value by Key in Key-Values")
        Set KeyValues = New Collection
        KeyValues.Add RestHelpers.CreateKeyValue("a", 123)
        KeyValues.Add RestHelpers.CreateKeyValue("b", 456)
        KeyValues.Add RestHelpers.CreateKeyValue("c", 789)
        
        .Expect(RestHelpers.FindInKeyValues(KeyValues, "b")).ToEqual 456
        .Expect(RestHelpers.FindInKeyValues(KeyValues, "d")).ToBeEmpty
    End With
    
    ' ============================================= '
    ' 5. Request preparation / handling
    ' ============================================= '
    
    ' FormatToMediaType
    With Specs.It("FormatToMediaType should handle custom converters")
        .Expect(RestHelpers.FormatToMediaType(WebFormat.custom, "custom-a")).ToEqual "X-a"
        .Expect(RestHelpers.FormatToMediaType(WebFormat.custom, "custom-b")).ToEqual "X-b"
        .Expect(RestHelpers.FormatToMediaType(WebFormat.custom, "custom-c")).ToEqual "X-c"
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
    ' TimeoutTimerExpired
    
    ' ============================================= '
    ' 7. Mac
    ' ============================================= '

#If Mac Then
    ' ExecuteInShell
    
    ' PrepareTextForShell
    ' --------------------------------------------- '
    With Specs.It("should prepare text for shell (wrap string and !)")
        .Expect(RestHelpers.PrepareTextForShell("""message""")).ToEqual """""message"""""
        .Expect(RestHelpers.PrepareTextForShell("!abc!")).ToEqual "'!'""abc""'!'"
    End With
#End If
    
    ' ============================================= '
    ' 8. Cryptography
    ' ============================================= '
    
    ' HMACSHA1
    ' --------------------------------------------- '
    With Specs.It("should calculate HMAC with SHA1 algorithm")
        .Expect(RestHelpers.HMACSHA1("test", "secret")).ToEqual "1aa349585ed7ecbd3b9c486a30067e395ca4b356"
        .Expect(RestHelpers.HMACSHA1("123456789", "987654321")).ToEqual "eea1a8e956b1b26067e6d0bef57e54490b8892a9"
    End With
    
    ' HMACSHA256
    ' --------------------------------------------- '
    With Specs.It("should calculate HMAC with SHA256 algorithm")
        .Expect(RestHelpers.HMACSHA256("test", "secret")).ToEqual "0329a06b62cd16b33eb6792be8c60b158d89a2ee3a876fce9a881ebb488c0914"
        .Expect(RestHelpers.HMACSHA256("123456789", "987654321")).ToEqual "3122584687113ac66d3c2f3c3518c789eef536a298121e0dbc82fc8fe7621e73"
    End With
    
    ' MD5
    ' --------------------------------------------- '
    With Specs.It("should MD5 hash string")
        .Expect(RestHelpers.MD5("test")).ToEqual "098f6bcd4621d373cade4e832627b4f6"
        .Expect(RestHelpers.MD5("123456789")).ToEqual "25f9e794323b453885f5181f1b624d0b"
        .Expect(RestHelpers.MD5("Mufasa:testrealm@host.com:Circle Of Life")).ToEqual "939e7578ed9e3c518a452acee763bce9"
    End With
    
    ' CreateNonce
    ' --------------------------------------------- '
    With Specs.It("should create Nonce of specified length")
        .Expect(Len(RestHelpers.CreateNonce)).ToEqual 32
        .Expect(Len(RestHelpers.CreateNonce(20))).ToEqual 20
    End With
    
    InlineRunner.RunSuite Specs
End Function

Function SimpleConverter(Body As Variant) As String
    Body.Add "response", "Goodbye!"
    SimpleConverter = RestHelpers.ConvertToJSON(Body)
End Function
Function SimpleParser(Content As String) As Object
    Set SimpleParser = RestHelpers.ParseJSON(Content)
    SimpleParser.Add "response", "Goodbye!"
End Function
Function ComplexParser(Body As Variant) As Object
    Dim Content As String
    Dim i As Integer
    
    For i = LBound(Body) To UBound(Body)
        Content = Content & Chr(Body(i))
    Next i
    
    Set ComplexParser = RestHelpers.ParseJSON("{""message"":""" & Content & """}")
    ComplexParser.Add "response", "Goodbye!"
End Function
