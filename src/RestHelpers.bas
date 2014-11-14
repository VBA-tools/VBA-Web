Attribute VB_Name = "RestHelpers"
''
' RestHelpers v4.0.0-beta.2
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Common helpers RestClient
'
' @dependencies: Microsoft Scripting Runtime
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Contents:
' 1. Logging
' 2. Converters and encoding
' 3. Url handling
' 4. Object/Dictionary/Collection/Array helpers
' 5. Request preparation / handling
' 6. Timing
' 7. Mac
' 8. Cryptography
' 9. Converters
' --------------------------------------------- '

#If Mac Then
Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As Long
Private Declare Function pclose Lib "libc.dylib" (ByVal File As Long) As Long
Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal Items As Long, ByVal stream As Long) As Long
Private Declare Function feof Lib "libc.dylib" (ByVal File As Long) As Long
#End If

#If Mac Then
#ElseIf Win64 Then
Private Declare PtrSafe Sub JSON_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (JSON_MemoryDestination As Any, JSON_MemorySource As Any, ByVal JSON_ByteLength As Long)
#Else
Private Declare Sub JSON_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (JSON_MemoryDestination As Any, JSON_MemorySource As Any, ByVal JSON_ByteLength As Long)
#End If

Public Type ShellResult
    Output As String
    ExitCode As Long
End Type

Private pDocumentHelper As Object
Private pElHelper As Object
Private pAsyncRequests As Dictionary

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum WebStatusCode
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
    RequestTimeout = 408
    UnsupportedMediaType = 415
    InternalServerError = 500
    BadGateway = 502
    ServiceUnavailable = 503
    GatewayTimeout = 504
End Enum
Public Enum WebMethod
    httpGET = 0
    httpPOST = 1
    httpPUT = 2
    httpDELETE = 3
    httpPATCH = 4
End Enum
Public Enum WebFormat
    plaintext = 0
    json = 1
    formurlencoded = 2
    xml = 3
    custom = 9
End Enum

Public EnableLogging As Boolean

' ============================================= '
' 1. Logging
' ============================================= '

''
' Log debug message with optional from description
'
' @param {String} Message
' @param {String} [From]
' --------------------------------------------- '
Public Sub LogDebug(Message As String, Optional From As String = "")
    If EnableLogging Then
        If From = "" Then
            From = "Excel-REST"
        End If
        
        Debug.Print From & ": " & Message
    End If
End Sub

''
' Log error message with optional from description and error number
'
' @param {String} Message
' @param {String} [From]
' @param {Long} [ErrNumber]
' --------------------------------------------- '
Public Sub LogError(Message As String, Optional From As String = "", Optional ErrNumber As Long = -1)
    If From = "" Then
        From = "Excel-REST"
    End If
    If ErrNumber >= 0 Then
        From = From & ": " & ErrNumber
    End If
    
    Debug.Print "ERROR - " & From & ": " & Message
End Sub

''
' Log request
'
' @param {RestRequest} Request
' --------------------------------------------- '
Public Sub LogRequest(Request As RestRequest, Optional FullUrl As String = "")
    If EnableLogging Then
        If FullUrl = "" Then
            FullUrl = Request.FormattedResource
        End If
    
        Debug.Print "--> Request - " & Format(Now, "Long Time")
        Debug.Print MethodToName(Request.Method) & " " & FullUrl
        
        Dim KeyValue As Dictionary
        For Each KeyValue In Request.Headers
            Debug.Print KeyValue("Key") & ": " & KeyValue("Value")
        Next KeyValue
        
        For Each KeyValue In Request.Cookies
            Debug.Print "Cookie: " & KeyValue("Key") & "=" & KeyValue("Value")
        Next KeyValue
        
        If Request.Body <> "" Then
            Debug.Print vbNewLine & Request.Body
        End If
        
        Debug.Print
    End If
End Sub

''
' Log response
'
' @param {RestResponse} Response
' --------------------------------------------- '
Public Sub LogResponse(Response As RestResponse, Request As RestRequest)
    If EnableLogging Then
        Debug.Print "<-- Response - " & Format(Now, "Long Time")
        Debug.Print Response.StatusCode & " " & Response.StatusDescription
        
        Dim KeyValue As Dictionary
        For Each KeyValue In Response.Headers
            Debug.Print KeyValue("Key") & ": " & KeyValue("Value")
        Next KeyValue
        
        For Each KeyValue In Response.Cookies
            Debug.Print "Cookie: " & KeyValue("Key") & "=" & KeyValue("Value")
        Next KeyValue
        
        Debug.Print vbNewLine & Response.Content & vbNewLine
    End If
End Sub

''
' Obfuscate message (for logging) by replacing with given character
'
' Example: ("Password", "#") -> ########
'
' @param {String} Secure
' @param {String} [Character = *]
' @return {String}
' --------------------------------------------- '
Public Function Obfuscate(Secure As String, Optional Character As String = "*") As String
    Obfuscate = String(Len(Secure), Character)
End Function

' ============================================= '
' 2. Converters and encoding
' ============================================= '

''
' Parse given JSON string into object (Dictionary or Collection)
'
' @param {String} JSON
' @return {Object}
' --------------------------------------------- '
' ParseJSON - Implemented in VBA-JSONConverter embedded below

''
' Convert object to JSON string
'
' @param {Variant} Obj
' @return {String}
' --------------------------------------------- '
' ConvertToJSON - Implemented in VBA-JSONConverter embedded below

''
' Parse url-encoded string to Dictionary
' TODO: Handle arrays and collections
'
' @param {String} UrlEncoded
' @return {Dictionary} Parsed
' --------------------------------------------- '
Public Function ParseUrlEncoded(Encoded As String) As Dictionary
    Dim Items As Variant
    Dim i As Integer
    Dim Parts As Variant
    Dim Parsed As New Dictionary
    Dim Key As String
    Dim Value As Variant
    
    Items = Split(Encoded, "&")
    For i = LBound(Items) To UBound(Items)
        Parts = Split(Items(i), "=")
        
        If UBound(Parts) - LBound(Parts) >= 1 Then
            ' TODO: Handle numbers, arrays, and object better here
            Key = UrlDecode(CStr(Parts(LBound(Parts))))
            Value = UrlDecode(CStr(Parts(LBound(Parts) + 1)))
            
            If Parsed.Exists(Key) Then
                Parsed(Key) = Value
            Else
                Parsed.Add Key, Value
            End If
        End If
    Next i
    
    Set ParseUrlEncoded = Parsed
End Function

''
' Convert dictionary/collection of key-value to url encoded string
'
' @param {Variant} Obj
' @return {String} UrlEncoded string (e.g. a=123&b=456&...)
' --------------------------------------------- '
Public Function ConvertToUrlEncoded(Obj As Variant) As String
    Dim Encoded As String

    If TypeOf Obj Is Collection Then
        Dim KeyValue As Dictionary
        For Each KeyValue In Obj
            If Len(Encoded) > 0 Then: Encoded = Encoded & "&"
            Encoded = Encoded & GetUrlEncodedKeyValue(KeyValue("Key"), KeyValue("Value"))
        Next KeyValue
    Else
        Dim Key As Variant
        For Each Key In Obj.Keys()
            If Len(Encoded) > 0 Then: Encoded = Encoded & "&"
            Encoded = Encoded & GetUrlEncodedKeyValue(Key, Obj(Key))
        Next Key
    End If
    
    ConvertToUrlEncoded = Encoded
End Function

''
' Parse XML string to XML
'
' @param {String} Encoded
' @return {Object} XML
' --------------------------------------------- '
Public Function ParseXML(Encoded As String) As Object
#If Mac Then
    LogError "ParseXML is not supported on Mac", "RestHelpers.ParseXML"
    Err.Raise vbObjectError + 1, "RestHelpers.ParseXML", "ParseXML is not supported on Mac"
#Else
    Set ParseXML = CreateObject("MSXML2.DOMDocument")
    ParseXML.Async = False
    ParseXML.LoadXML Encoded
#End If
End Function

''
' Convert MSXML2.DomDocument to string
'
' @param {Object: MSXML2.DomDocument} XML
' @return {String} XML string
' --------------------------------------------- '

Public Function ConvertToXML(Obj As Variant) As String
    On Error Resume Next
    ConvertToXML = Trim(Replace(Obj.xml, vbCrLf, ""))
End Function

''
' Parse given string into object (Dictionary or Collection) for given format
'
' @param {String} Value
' @param {WebFormat} Format
' @return {Object}
' --------------------------------------------- '
Public Function ParseByFormat(Value As String, Format As WebFormat) As Object
    Select Case Format
    Case WebFormat.json
        Set ParseByFormat = ParseJSON(Value)
    Case WebFormat.formurlencoded
        Set ParseByFormat = ParseUrlEncoded(Value)
    Case WebFormat.xml
        Set ParseByFormat = ParseXML(Value)
    End Select
End Function

''
' Convert object to given format
'
' @param {Variant} Obj
' @param {WebFormat} Format
' @return {String}
' --------------------------------------------- '
Public Function ConvertToFormat(Obj As Variant, Format As WebFormat) As String
    Select Case Format
    Case WebFormat.json
        ConvertToFormat = ConvertToJSON(Obj)
    Case WebFormat.formurlencoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case WebFormat.xml
        ConvertToFormat = ConvertToXML(Obj)
    End Select
End Function

''
' Url encode the given string
' Reference: http://www.blooberry.com/indexdot/html/topics/urlencoding.htm
'
' @param {Variant} Text The raw string to encode
' @param {Boolean} [SpaceAsPlus = False] Use plus sign for encoded spaces (otherwise %20)
' @param {Boolean} [EncodeUnsafe = True] Encode unsafe characters
' @return {String} Encoded string
' --------------------------------------------- '
Public Function UrlEncode(Text As Variant, Optional SpaceAsPlus As Boolean = False, Optional EncodeUnsafe As Boolean = True) As String
    Dim UrlVal As String
    Dim StringLen As Long
    
    UrlVal = CStr(Text)
    StringLen = Len(UrlVal)
    
    If StringLen > 0 Then
        ReDim Result(StringLen) As String
        Dim i As Long
        Dim CharCode As Integer
        Dim Char As String
        Dim Space As String
        
        ' Set space value
        If SpaceAsPlus Then
            Space = "+"
        Else
            Space = "%20"
        End If
        
        ' Loop through string characters
        For i = 1 To StringLen
            ' Get character and ascii code
            Char = Mid$(UrlVal, i, 1)
            CharCode = asc(Char)
            
            Select Case CharCode
                Case 36, 38, 43, 44, 47, 58, 59, 61, 63, 64
                    ' Reserved characters
                    Result(i) = "%" & Hex(CharCode)
                Case 32, 34, 35, 37, 60, 62, 91 To 94, 96, 123 To 126
                    ' Unsafe characters
                    If EncodeUnsafe Then
                        If CharCode = 32 Then
                            Result(i) = Space
                        Else
                            Result(i) = "%" & Hex(CharCode)
                        End If
                    End If
                Case Else
                    Result(i) = Char
            End Select
        Next i
        UrlEncode = Join(Result, "")
    End If
End Function

''
' Url decode the given encoded string
'
' @param {String} Encoded
' @return {String} Decoded string
' --------------------------------------------- '
Public Function UrlDecode(Encoded As String) As String
    Dim StringLen As Long
    StringLen = Len(Encoded)
    
    If StringLen > 0 Then
        Dim i As Long
        Dim Result As String
        Dim Temp As String
        
        For i = 1 To StringLen
            Temp = Mid$(Encoded, i, 1)
            
            If Temp = "+" Then
                Temp = " "
            ElseIf Temp = "%" And StringLen >= i + 2 Then
                Temp = Mid$(Encoded, i + 1, 2)
                Temp = Chr(CInt("&H" & Temp))
                
                i = i + 2
            End If
                
            Result = Result & Temp
        Next i
        
        UrlDecode = Result
    End If
End Function

''
' Url encode the given string
'
' @param {Variant} Text The raw string to encode
' @return {String} Encoded string
' --------------------------------------------- '
Public Function Base64Encode(Text As String) As String
    Base64Encode = Replace(StringToBase64(Text), vbLf, "")
End Function

' ============================================= '
' 3. Url handling
' ============================================= '

''
' Join Url with /
'
' @param {String} LeftSide
' @param {String} RightSide
' @return {String} Joined url
' --------------------------------------------- '
Public Function JoinUrl(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "/" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "/" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If
    
    If LeftSide <> "" And RightSide <> "" Then
        JoinUrl = LeftSide & "/" & RightSide
    Else
        JoinUrl = LeftSide & RightSide
    End If
End Function

''
' Get Url parts
'
' Example:
' "https://www.google.com/a/b/c.html?a=1&b=2#hash" ->
' - Protocol = https
' - Host = www.google.com
' - Port = 443
' - Path = /a/b/c.html
' - Querystring = a=1&b=2
' - Hash = hash
'
' "https://localhost:3000/a/b/c.html?a=1&b=2#hash" ->
' - Protocol = https
' - Host = localhost
' - Port = 3000
' - Path = /a/b/c.html
' - Querystring = a=1&b=2
' - Hash = hash
'
' @param {String} Url
' @return {Dictionary} Parts of url
' Protocol, Host, Hostname, Port, Uri, Querystring, Hash
' --------------------------------------------- '
Public Function UrlParts(Url As String) As Dictionary
    Dim Parts As New Dictionary
    
#If Mac Then
    ' Run perl script to parse url
    ' Add Protocol if missing
    Dim AddedProtocol As Boolean
    If InStr(1, Url, "://") <= 0 Then
        AddedProtocol = True
        If InStr(1, Url, "//") = 1 Then
            Url = "http" & Url
        Else
            Url = "http://" & Url
        End If
    End If
    
    Dim Command As String
    Dim Result As ShellResult
    Dim Results As Variant
    Dim ResultPart As Variant
    Dim EqualsIndex As Long
    Dim Key As String
    Dim Value As String
    Command = "perl -e '{use URI::URL;" & vbNewLine & _
        "$url = new URI::URL """ & Url & """;" & vbNewLine & _
        "print ""Protocol="" . $url->scheme;" & vbNewLine & _
        "print "" | Host="" . $url->host;" & vbNewLine & _
        "print "" | Port="" . $url->port;" & vbNewLine & _
        "print "" | FullPath="" . $url->full_path;" & vbNewLine & _
        "print "" | Hash="" . $url->frag;" & vbNewLine & _
    "}'"

    Results = Split(ExecuteInShell(Command).Output, " | ")
    For Each ResultPart In Results
        EqualsIndex = InStr(1, ResultPart, "=")
        Key = Trim(VBA.Mid$(ResultPart, 1, EqualsIndex - 1))
        Value = Trim(VBA.Mid$(ResultPart, EqualsIndex + 1))
        
        If Key = "FullPath" Then
            ' For properly escaped path and querystring, need to use full_path
            ' But, need to split FullPath into Path...?Querystring
            Dim QueryIndex As Integer
            
            QueryIndex = InStr(1, Value, "?")
            If QueryIndex > 0 Then
                Parts.Add "Path", Mid$(Value, 1, QueryIndex - 1)
                Parts.Add "Querystring", Mid$(Value, QueryIndex + 1)
            Else
                Parts.Add "Path", Value
                Parts.Add "Querystring", ""
            End If
        Else
            Parts.Add Key, Value
        End If
    Next ResultPart
    
    If AddedProtocol And Parts.Exists("Protocol") Then
        Parts.Remove "Protocol"
    End If
#Else
    ' Create document/element is expensive, cache after creation
    If pDocumentHelper Is Nothing Or pElHelper Is Nothing Then
        Set pDocumentHelper = CreateObject("htmlfile")
        Set pElHelper = pDocumentHelper.createElement("a")
    End If
    
    pElHelper.href = Url
    Parts.Add "Protocol", Replace(pElHelper.Protocol, ":", "", Count:=1)
    Parts.Add "Host", pElHelper.hostname
    Parts.Add "Port", pElHelper.port
    Parts.Add "Path", pElHelper.pathname
    Parts.Add "Querystring", Replace(pElHelper.Search, "?", "", Count:=1)
    Parts.Add "Hash", Replace(pElHelper.Hash, "#", "", Count:=1)
#End If

    If Parts("Protocol") = "localhost" Then
        ' localhost:port/... was passed in without protocol
        Dim PathParts As Variant
        PathParts = Split(Parts("Path"), "/")
        
        Parts("Port") = PathParts(0)
        Parts("Protocol") = ""
        Parts("Host") = "localhost"
        Parts("Path") = Replace(Parts("Path"), Parts("Port"), "", Count:=1)
    End If
    If Left(Parts("Path"), 1) <> "/" Then
        Parts("Path") = "/" & Parts("Path")
    End If

    Set UrlParts = Parts
End Function

' ============================================= '
' 4. Object/Dictionary/Collection/Array helpers
' ============================================= '

''
' Combine two dictionaries, folding the second into the first
'
' @param {Dictionary} OriginalObj dictionary to add values to
' @param {Dictionary} NewObj New object containing values to add to original object
' @param {Boolean} [OverwriteOriginal=True] Overwrite any values that already exist in the original object
' @return {Dictionary} Combined object
' --------------------------------------------- '
Public Function CombineDictionaries(ByVal OriginalObj As Dictionary, ByVal NewObj As Dictionary, _
    Optional OverwriteOriginal As Boolean = True) As Dictionary
    
    Dim Combined As New Dictionary
    
    Dim OriginalKey As Variant
    Dim Key As Variant
    
    If Not OriginalObj Is Nothing Then
        For Each Key In OriginalObj.Keys()
            Combined.Add Key, OriginalObj(Key)
        Next Key
    End If
    If Not NewObj Is Nothing Then
        For Each Key In NewObj.Keys()
            If Combined.Exists(Key) And OverwriteOriginal Then
                Combined(Key) = NewObj(Key)
            ElseIf Not Combined.Exists(Key) Then
                Combined.Add Key, NewObj(Key)
            End If
        Next Key
    End If
    
    Set CombineDictionaries = Combined
End Function

''
' Apply whitelist to given object to filter out unwanted key/values
'
' @param {Dictionary} Original model to filter
' @param {Variant} WhiteList Array|String of value(s) to retain in the model
' @return {Dictionary} Filtered object
' --------------------------------------------- '
Public Function FilterDictionary(ByVal Original As Dictionary, Whitelist As Variant) As Dictionary
    Dim Filtered As New Dictionary
    Dim i As Integer
    
    If IsArray(Whitelist) Then
        For i = LBound(Whitelist) To UBound(Whitelist)
            If Original.Exists(Whitelist(i)) Then
                Filtered.Add Whitelist(i), Original(Whitelist(i))
            End If
        Next i
    ElseIf VarType(Whitelist) = vbString Then
        If Original.Exists(Whitelist) Then
            Filtered.Add Whitelist, Original(Whitelist)
        End If
    End If
    
    Set FilterDictionary = Filtered
End Function

''
' Check if given is an array
'
' @param {Object} Obj
' @return {Boolean}
' --------------------------------------------- '
Public Function IsArray(Obj As Variant) As Boolean
    Select Case VarType(Obj)
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        IsArray = True
    End Select
End Function

''
' Clone dictionary
'
' @param {Dictionary} Dict
' @return {Dictionary}
' --------------------------------------------- '
Public Function CloneDictionary(Dict As Dictionary) As Dictionary
    Set CloneDictionary = New Dictionary
    Dim Key As Variant
    For Each Key In Dict.Keys
        CloneDictionary.Add CStr(Key), Dict(Key)
    Next Key
End Function

''
' Clone collection
'
' Note: Keys are not transferred to clone
'
' @param {Collection} Coll
' @return {Collection}
' --------------------------------------------- '
Public Function CloneCollection(Coll As Collection) As Collection
    Set CloneCollection = New Collection
    Dim Item As Variant
    For Each Item In Coll
        CloneCollection.Add Item
    Next Item
End Function

''
' Helper for creating key-value Dictionary for collection
'
' @param {String} Key
' @param {Variant} Value
' @return {Dictionary}
' --------------------------------------------- '
Public Function CreateKeyValue(Key As String, Value As Variant) As Dictionary
    Dim KeyValue As New Dictionary
    KeyValue("Key") = Key
    KeyValue("Value") = Value
    Set CreateKeyValue = KeyValue
End Function

''
' Helper for finding key-value in Collection of key-value
'
' @param {Collection} KeyValues
' @param {String} Key to find
' @return {Variant}
' --------------------------------------------- '
Public Function FindInKeyValues(KeyValues As Collection, Key As Variant) As Variant
    Dim KeyValue As Dictionary
    For Each KeyValue In KeyValues
        If KeyValue("Key") = Key Then
            FindInKeyValues = KeyValue("Value")
            Exit Function
        End If
    Next KeyValue
End Function

' ============================================= '
' 5. Request preparation / handling
' ============================================= '

''
' Set headers to http object for given request
'
' @param {WinHttpRequest} Http request
' @param {RestRequest} Request
' --------------------------------------------- '
Public Sub SetHeadersForHttp(ByRef Http As Object, Request As RestRequest)
    Dim HeaderKeyValue As Dictionary
    For Each HeaderKeyValue In Request.Headers
        Http.setRequestHeader HeaderKeyValue("Key"), HeaderKeyValue("Value")
    Next HeaderKeyValue
    
    Dim CookieKeyValue As Dictionary
    For Each CookieKeyValue In Request.Cookies
        Http.setRequestHeader "Cookie", CookieKeyValue("Key") & "=" & CookieKeyValue("Value")
    Next CookieKeyValue
End Sub

''
' Create simple response
'
' @param {WebStatusCode} StatusCode
' @param {String} StatusDescription
' @return {RestResponse}
' --------------------------------------------- '
Public Function CreateResponse(StatusCode As WebStatusCode, StatusDescription As String) As RestResponse
    Set CreateResponse = New RestResponse
    CreateResponse.StatusCode = StatusCode
    CreateResponse.StatusDescription = StatusDescription
End Function

''
' Create response for http
'
' @param {WinHttpRequest} Http
' @param {WebFormat} [Format=json]
' @return {RestResponse}
' --------------------------------------------- '
Public Function CreateResponseFromHttp(ByRef Http As Object, Optional Format As WebFormat = WebFormat.json) As RestResponse
    Set CreateResponseFromHttp = New RestResponse
    
    CreateResponseFromHttp.StatusCode = Http.Status
    CreateResponseFromHttp.StatusDescription = Http.StatusText
    CreateResponseFromHttp.Body = Http.ResponseBody
    CreateResponseFromHttp.Content = Http.ResponseText
    
    ' Convert content to data by format
    If Format <> WebFormat.plaintext Then
        On Error Resume Next
        Set CreateResponseFromHttp.Data = RestHelpers.ParseByFormat(Http.ResponseText, Format)
        On Error GoTo 0
    End If
    
    ' Extract headers
    Set CreateResponseFromHttp.Headers = ExtractHeaders(Http.getAllResponseHeaders)
    
    ' Extract cookies
    Set CreateResponseFromHttp.Cookies = ExtractCookies(CreateResponseFromHttp.Headers)
End Function

''
' Create response for cURL
' References:
' http://www.w3.org/Protocols/rfc2616/rfc2616-sec6.html
' http://curl.haxx.se/libcurl/c/libcurl-errors.html
'
' @param {String} Raw result from cURL
' @return {RestResponse}
' --------------------------------------------- '
Public Function CreateResponseFromCURL(Result As ShellResult, Optional Format As WebFormat = WebFormat.json) As RestResponse
    Dim StatusCode As Long
    Dim StatusText As String
    Dim Headers As String
    Dim Body As Variant
    Dim ResponseText As String
    
    If Result.ExitCode > 0 Then
        Dim ErrorNumber As Long
        
        ErrorNumber = Result.ExitCode / 256
        ' 5 - CURLE_COULDNT_RESOLVE
        ' 7 - CURLE_COULDNT_CONNECT
        ' 28 - CURLE_OPERATION_TIMEDOUT
        If ErrorNumber = 5 Or ErrorNumber = 7 Or ErrorNumber = 28 Then
            Set CreateResponseFromCURL = CreateResponse(WebStatusCode.RequestTimeout, "Request Timeout")
        Else
            LogError "cURL Error: " & ErrorNumber, "RestHelpers.CreateResponseFromCURL"
        End If
        
        Exit Function
    End If
    
    Dim Lines() As String
    Lines = Split(Result.Output, vbCrLf)
    
    ' Extract status code and text from status line
    Dim StatusLine As String
    Dim StatusLineParts() As String
    StatusLine = Lines(0)
    StatusLineParts = Split(StatusLine)
    StatusCode = CLng(StatusLineParts(1))
    StatusText = Mid$(StatusLine, InStr(1, StatusLine, StatusCode) + 4)
    
    ' Find blank line before body
    Dim Line As Variant
    Dim BlankLineIndex
    BlankLineIndex = 0
    For Each Line In Lines
        If Trim(Line) = "" Then
            Exit For
        End If
        BlankLineIndex = BlankLineIndex + 1
    Next Line
    
    ' Extract body and headers strings
    Dim HeaderLines() As String
    Dim BodyLines() As String
    Dim ReadIndex As Long
    Dim WriteIndex As Long
    
    ReDim HeaderLines(0 To BlankLineIndex - 2)
    ReDim BodyLines(0 To UBound(Lines) - BlankLineIndex - 1)
    
    WriteIndex = 0
    For ReadIndex = 1 To BlankLineIndex - 1
        HeaderLines(WriteIndex) = Lines(ReadIndex)
        WriteIndex = WriteIndex + 1
    Next ReadIndex
    
    WriteIndex = 0
    For ReadIndex = BlankLineIndex + 1 To UBound(Lines)
        BodyLines(WriteIndex) = Lines(ReadIndex)
        WriteIndex = WriteIndex + 1
    Next ReadIndex
    
    ResponseText = Join$(BodyLines, vbCrLf)
    Body = StringToANSIBytes(ResponseText)
    
    ' Create Response
    Set CreateResponseFromCURL = New RestResponse
    CreateResponseFromCURL.StatusCode = StatusCode
    CreateResponseFromCURL.StatusDescription = StatusText
    CreateResponseFromCURL.Body = Body
    CreateResponseFromCURL.Content = ResponseText
    
    ' Convert content to data by format
    If Format <> WebFormat.plaintext Then
        On Error Resume Next
        Set CreateResponseFromCURL.Data = RestHelpers.ParseByFormat(CreateResponseFromCURL.Content, Format)
        On Error GoTo 0
    End If
    
    ' Extract headers
    Set CreateResponseFromCURL.Headers = ExtractHeaders(Join$(HeaderLines, vbCrLf))
    
    ' Extract cookies
    Set CreateResponseFromCURL.Cookies = ExtractCookies(CreateResponseFromCURL.Headers)
End Function

''
' Extract headers from response headers
'
' @param {String} ResponseHeaders
' @return {Collection} Headers
' --------------------------------------------- '
Public Function ExtractHeaders(ResponseHeaders As String) As Collection
    Dim Headers As New Collection
    Dim Header As Dictionary
    Dim Multiline As Boolean
    Dim Key As String
    Dim Value As String
    
    Dim Lines As Variant
    Lines = Split(ResponseHeaders, vbCrLf)
    
    Dim i As Integer
    For i = LBound(Lines) To (UBound(Lines) + 1)
        If i > UBound(Lines) Then
            Headers.Add Header
        ElseIf Lines(i) <> "" Then
            If InStr(1, Lines(i), ":") = 0 And Not Header Is Nothing Then
                ' Assume part of multi-line header
                Multiline = True
            ElseIf Multiline Then
                ' Close out multi-line string
                Multiline = False
                Headers.Add Header
            ElseIf Not Header Is Nothing Then
                Headers.Add Header
            End If
            
            If Not Multiline Then
                Set Header = CreateKeyValue( _
                    Key:=Trim(Mid$(Lines(i), 1, InStr(1, Lines(i), ":") - 1)), _
                    Value:=Trim(Mid$(Lines(i), InStr(1, Lines(i), ":") + 1, Len(Lines(i)))) _
                )
            Else
                Header("Value") = Header("Value") & vbCrLf & Lines(i)
            End If
        End If
    Next i
    
    Set ExtractHeaders = Headers
End Function

''
' Extract cookies from response headers
'
' @param {Collection} Headers
' @return {Collection} Cookies
' --------------------------------------------- '
Public Function ExtractCookies(Headers As Collection) As Collection
    Dim Cookies As New Collection
    Dim Cookie As String
    Dim Key As String
    Dim Value As String
    Dim Header As Dictionary
    
    For Each Header In Headers
        If Header("Key") = "Set-Cookie" Then
            Cookie = Header("Value")
            Key = Mid$(Cookie, 1, InStr(1, Cookie, "=") - 1)
            Value = Mid$(Cookie, InStr(1, Cookie, "=") + 1, Len(Cookie))
            
            If InStr(1, Value, ";") Then
                Value = Mid$(Value, 1, InStr(1, Value, ";") - 1)
            End If
            
            Cookies.Add CreateKeyValue(Key, UrlDecode(Value))
        End If
    Next Header
    
    Set ExtractCookies = Cookies
End Function

''
' Create request from options
'
' @param {Dictionary} Options
' - Headers
' - Cookies
' - QuerystringParams
' - UrlSegments
' --------------------------------------------- '
Public Function CreateRequestFromOptions(Options As Dictionary) As RestRequest
    Dim Request As New RestRequest
    
    If Not IsEmpty(Options) And Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set Request.Headers = Options("Headers")
        End If
        If Options.Exists("Cookies") Then
            Set Request.Cookies = Options("Cookies")
        End If
        If Options.Exists("QuerystringParams") Then
            Set Request.QuerystringParams = Options("QuerystringParams")
        End If
        If Options.Exists("UrlSegments") Then
            Set Request.UrlSegments = Options("UrlSegments")
        End If
    End If
    
    Set CreateRequestFromOptions = Request
End Function

''
' Update response with another response
'
' @param {RestResponse) Original (Updated by reference)
' @param {RestResponse) Updated
' @return {RestResponse}
' --------------------------------------------- '
Public Function UpdateResponse(ByRef Original As RestResponse, Updated As RestResponse) As RestResponse
    Original.StatusCode = Updated.StatusCode
    Original.StatusDescription = Updated.StatusDescription
    Original.Content = Updated.Content
    Original.Body = Updated.Body
    Set Original.Headers = Updated.Headers
    Set Original.Cookies = Updated.Cookies
    
    If Not IsEmpty(Updated.Data) Then
        If IsObject(Updated.Data) Then
            Set Original.Data = Updated.Data
        Else
            Original.Data = Updated.Data
        End If
    End If
    
    Set UpdateResponse = Original
End Function

''
' Get content-type for format
'
' @param {WebFormat} Format
' @return {String}
' --------------------------------------------- '
Public Function FormatToContentType(Format As WebFormat) As String
    Select Case Format
    Case WebFormat.formurlencoded
        FormatToContentType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case WebFormat.json
        FormatToContentType = "application/json"
    Case WebFormat.xml
        FormatToContentType = "application/xml"
    Case WebFormat.plaintext
        FormatToContentType = "text/plain"
    End Select
End Function

''
' Get name for method
'
' @param {} Method
' @return {String}
' --------------------------------------------- '
Public Function MethodToName(Method As WebMethod) As String
    Select Case Method
    Case WebMethod.httpDELETE
        MethodToName = "DELETE"
    Case WebMethod.httpPUT
        MethodToName = "PUT"
    Case WebMethod.httpPATCH
        MethodToName = "PATCH"
    Case WebMethod.httpPOST
        MethodToName = "POST"
    Case WebMethod.httpGET
        MethodToName = "GET"
    End Select
End Function

''
' Add request to watched requests
'
' @param {RestAsyncWrapper} AsyncWrapper
' --------------------------------------------- '
Public Sub AddAsyncRequest(AsyncWrapper As Object)
    If pAsyncRequests Is Nothing Then: Set pAsyncRequests = New Dictionary
    If Not AsyncWrapper.Request Is Nothing Then
        pAsyncRequests.Add AsyncWrapper.Request.Id, AsyncWrapper
    End If
End Sub

''
' Get watched request
'
' @param {String} RequestId
' @return {RestAsyncWrapper}
' --------------------------------------------- '
Public Function GetAsyncRequest(RequestId As String) As Object
    If pAsyncRequests.Exists(RequestId) Then
        Set GetAsyncRequest = pAsyncRequests(RequestId)
    End If
End Function

''
' Remove request from watched requests
'
' @param {String} RequestId
' --------------------------------------------- '
Public Sub RemoveAsyncRequest(RequestId As String)
    If Not pAsyncRequests Is Nothing Then
        If pAsyncRequests.Exists(RequestId) Then: pAsyncRequests.Remove RequestId
    End If
End Sub

' ============================================= '
' 6. Timing
' ============================================= '

''
' Start timeout timer for request
'
' @param {RestRequest} Request
' @param {Long} TimeoutMS
' --------------------------------------------- '
Public Sub StartTimeoutTimer(AsyncWrapper As Object, TimeoutMS As Long)
    ' Round ms to seconds with minimum of 1 second if ms > 0
    Dim TimeoutS As Long
    TimeoutS = Round(TimeoutMS / 1000, 0)
    If TimeoutMS > 0 And TimeoutS = 0 Then
        TimeoutS = 1
    End If

    AddAsyncRequest AsyncWrapper
    Application.OnTime Now + TimeValue("00:00:" & TimeoutS), "'RestHelpers.TimeoutTimerExpired """ & AsyncWrapper.Request.Id & """'"
End Sub

''
' Stop timeout timer for request
'
' @param {RestRequest} Request
' --------------------------------------------- '
Public Sub StopTimeoutTimer(AsyncWrapper As Object)
    If Not AsyncWrapper.Request Is Nothing Then
        RemoveAsyncRequest AsyncWrapper.Request.Id
    End If
End Sub

''
' Handle timeout timers expiring
'
' @param {String} RequestId
' --------------------------------------------- '
Public Sub TimeoutTimerExpired(RequestId As String)
    Dim AsyncWrapper As Object
    Set AsyncWrapper = GetAsyncRequest(RequestId)
    
    If Not AsyncWrapper Is Nothing Then
        StopTimeoutTimer AsyncWrapper
        
        LogDebug "Async Timeout: " & AsyncWrapper.Request.FormattedResource, "RestHelpers.TimeoutTimerExpired"
        AsyncWrapper.TimedOut
    End If
End Sub

' ============================================= '
' 7. Mac
' ============================================= '
#If Mac Then

''
' Execute the given command
'
' @param {String} Command
' @return {ShellResult}
' --------------------------------------------- '
Public Function ExecuteInShell(Command As String) As ShellResult
    Dim File As Long
    Dim Chunk As String
    Dim Read As Long
    
    On Error GoTo ErrorHandling
    File = popen(Command, "r")
    
    If File = 0 Then
        ' TODO
        Exit Function
    End If
    
    Do While feof(File) = 0
        Chunk = VBA.Space$(50)
        Read = fread(Chunk, 1, Len(Chunk) - 1, File)
        If Read > 0 Then
            Chunk = VBA.Left$(Chunk, Read)
            ExecuteInShell.Output = ExecuteInShell.Output & Chunk
        End If
    Loop

ErrorHandling:
    ExecuteInShell.ExitCode = pclose(File)
End Function

''
' Prepare text for shell
' Wrap in "..." and replace ! with '!' (reserved in bash)
'
' @param {String} Text
' @return {String}
' --------------------------------------------- '
Public Function PrepareTextForShell(ByVal Text As String) As String
    Text = Replace("""" & Text & """", "!", """'!'""")
    
    ' Guard for ! at beginning or end ("'!'"..." or "..."'!'")
    If Left(Text, 3) = """""'" Then
        Text = Right(Text, Len(Text) - 2)
    End If
    If Right(Text, 3) = "'""""" Then
        Text = Left(Text, Len(Text) - 2)
    End If
    
    PrepareTextForShell = Text
End Function

#End If

' ============================================= '
' 8. Cryptography
' ============================================= '

''
' Perform HMAC-SHA1 on string and return as Hex or Base64
' [Does VBA have a Hash_HMAC](http://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac)
'
' @param {String} Text
' @param {String} Secret
' @param {String} [Format = Hex] Hex or Base64
' @return {String} HMAC-SHA1
' --------------------------------------------- '
Public Function HMACSHA1(Text As String, Secret As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -sha1 -hmac " & PrepareTextForShell(Secret)
    HMACSHA1 = Replace(ExecuteInShell(Command).Output, vbLf, "")
#Else
    Dim Crypto As Object
    Dim Bytes() As Byte
    
    Set Crypto = CreateObject("System.Security.Cryptography.HMACSHA1")
    Crypto.Key = StringToANSIBytes(Secret)
    Bytes = Crypto.ComputeHash_2(StringToANSIBytes(Text))
    HMACSHA1 = ANSIBytesToHex(Bytes)
#End If

    If Format = "Base64" Then
        HMACSHA1 = HexToBase64(HMACSHA1)
    End If
End Function

''
' Perform HMAC-SHA256 on string and return as Hex or Base64
'
' @param {String} Text
' @param {String} Secret
' @param {String} [Format = Hex] Hex or Base64
' @return {String} HMAC-SHA256
' --------------------------------------------- '
Public Function HMACSHA256(Text As String, Secret As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -sha256 -hmac " & PrepareTextForShell(Secret)
    HMACSHA256 = Replace(ExecuteInShell(Command).Output, vbLf, "")
#Else
    Dim Crypto As Object
    Dim Bytes() As Byte
    
    Set Crypto = CreateObject("System.Security.Cryptography.HMACSHA256")
    Crypto.Key = StringToANSIBytes(Secret)
    Bytes = Crypto.ComputeHash_2(StringToBytes(Text))
    HMACSHA256 = ANSIBytesToHex(Bytes)
#End If

    If Format = "Base64" Then
        HMACSHA256 = HexToBase64(HMACSHA256)
    End If
End Function

''
' Perform MD5 Hash on string and return as Hex or Base64
' Source: http://www.di-mgt.com.au/src/basMD5.bas.html
'
' @param {String} Text
' @param {String} [Format = Hex] Hex or Base64
' @return {String} MD5 Hash
' --------------------------------------------- '
Public Function MD5(Text As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -md5"
    MD5 = Replace(ExecuteInShell(Command).Output, vbLf, "")
#Else
    Dim Crypto As Object
    Dim Bytes() As Byte
    
    Set Crypto = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    Bytes = Crypto.ComputeHash_2(StringToANSIBytes(Text))
    MD5 = ANSIBytesToHex(Bytes)
#End If

    If Format = "Base64" Then
        MD5 = HexToBase64(MD5)
    End If
End Function

''
' Create random alphanumeric nonce
'
' @param {Integer} [NonceLength=32]
' @return {String} Randomly generated nonce
' --------------------------------------------- '
Public Function CreateNonce(Optional NonceLength As Integer = 32) As String
    Dim Str As String
    Dim Count As Integer
    Dim Result As String
    Dim random As Integer
    
    Str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    Result = ""
    
    For Count = 1 To NonceLength
        random = Int(((Len(Str) - 1) * Rnd) + 1)
        Result = Result + Mid$(Str, random, 1)
    Next
    CreateNonce = Result
End Function

Private Function StringToANSIBytes(Text As String) As Byte()
    Dim Bytes() As Byte
    Dim ANSIBytes() As Byte
    Dim ByteIndex As Long
    Dim ANSIIndex As Long
    
    If Len(Text) > 0 Then
        ' Take first byte from unicode bytes
        Bytes = Text
        ReDim ANSIBytes(Int(UBound(Bytes) / 2))
        
        ANSIIndex = LBound(Bytes)
        For ByteIndex = LBound(Bytes) To UBound(Bytes) Step 2
            ANSIBytes(ANSIIndex) = Bytes(ByteIndex)
            ANSIIndex = ANSIIndex + 1
        Next ByteIndex
    End If
    
    StringToANSIBytes = ANSIBytes
End Function

Private Function ANSIBytesToString(Bytes() As Byte) As String
    Dim i As Long
    For i = LBound(Bytes) To UBound(Bytes)
        ANSIBytesToString = ANSIBytesToString & VBA.Chr$(Bytes(i))
    Next i
End Function

Private Function HexToANSIBytes(Hex As String) As Byte()
    Dim Bytes() As Byte
    Dim HexIndex As Integer
    Dim ByteIndex As Integer

    ' Remove linefeeds
    Hex = VBA.Replace(Hex, vbLf, "")

    ReDim Bytes(VBA.Len(Hex) / 2 - 1)
    ByteIndex = 0
    For HexIndex = 1 To Len(Hex) Step 2
        Bytes(ByteIndex) = VBA.CLng("&H" & VBA.Mid$(Hex, HexIndex, 2))
        ByteIndex = ByteIndex + 1
    Next HexIndex
    
    HexToANSIBytes = Bytes
End Function

Private Function ANSIBytesToHex(Bytes() As Byte)
    Dim i As Long
    For i = LBound(Bytes) To UBound(Bytes)
        ANSIBytesToHex = ANSIBytesToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(Bytes(i)), 2))
    Next i
End Function

Private Function StringToHex(ByVal Text As String) As String
    ' Convert single-byte character to hex
    ' (May need better international handling in the future)
    Dim Bytes() As Byte
    Dim i As Integer
    
    Bytes = StringToANSIBytes(Text)
    For i = LBound(Bytes) To UBound(Bytes)
        StringToHex = StringToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(Bytes(i)), 2))
    Next i
End Function

Private Function StringToBase64(ByVal Text As String) As String
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl base64"
    StringToBase64 = ExecuteInShell(Command).Output
#Else
    ' Use XML to convert to Base64
    ' but XML requires bytes, so convert to bytes first
    Dim xml As Object
    Dim Node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    Set Node = xml.createElement("b64")
    Node.DataType = "bin.base64"
    Node.nodeTypedValue = StringToANSIBytes(Text)
    StringToBase64 = Node.Text

    Set Node = Nothing
    Set xml = Nothing
#End If
End Function

Private Function HexToBase64(ByVal Hex As String) As String
    HexToBase64 = StringToBase64(ANSIBytesToString(HexToANSIBytes(Hex)))
End Function

' ============================================= '
' 9. Converters
' ============================================= '

''
' Helper for url-encoded to create key=value pair
'
' @param {Variant} Key
' @param {Variant} Value
' @return {String}
' --------------------------------------------- '
Private Function GetUrlEncodedKeyValue(Key As Variant, Value As Variant) As String
    ' Convert boolean to lowercase
    If VarType(Value) = VBA.vbBoolean Then
        If Value Then
            Value = "true"
        Else
            Value = "false"
        End If
    End If
    
    ' Url encode key and value (using + for spaces)
    GetUrlEncodedKeyValue = UrlEncode(Key, SpaceAsPlus:=True) & "=" & UrlEncode(Value, SpaceAsPlus:=True)
End Function

''
' VBA-JSONConverter v1.0.0-beta.1
' (c) Tim Hall - https://github.com/timhall/VBA-JSONConverter
'
' JSON Converter for VBA
'
' Errors (513-65535 available):
' 10001 - JSON parse error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Moved to top)
'#If Mac Then
'#ElseIf Win64 Then
'Private Declare PtrSafe Sub JSON_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'    (JSON_MemoryDestination As Any, JSON_MemorySource As Any, ByVal JSON_ByteLength As Long)
'#Else
'Private Declare Sub JSON_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'    (JSON_MemoryDestination As Any, JSON_MemorySource As Any, ByVal JSON_ByteLength As Long)
'#End If

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @param {String} JSON_String
' @return {Object} (Dictionary or Collection)
' -------------------------------------- '
Public Function ParseJSON(ByVal JSON_String As String, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Object
    Dim JSON_Index As Long
    JSON_Index = 1
    
    ' Remove vbCr, vbLf, and vbTab from JSON_String
    JSON_String = VBA.Replace(VBA.Replace(VBA.Replace(JSON_String, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case VBA.Mid$(JSON_String, JSON_Index, 1)
    Case "{"
        Set ParseJSON = JSON_ParseObject(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
    Case "["
        Set ParseJSON = JSON_ParseArray(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @param {Variant} JSON_DictionaryCollectionOrArray (Dictionary, Collection, or Array)
' @return {String}
' -------------------------------------- '
Public Function ConvertToJSON(ByVal JSON_DictionaryCollectionOrArray As Variant, Optional JSON_ConvertLargeNumbersFromString As Boolean = True) As String
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    Dim JSON_Index As Long
    Dim JSON_LBound As Long
    Dim JSON_UBound As Long
    Dim JSON_IsFirstItem As Boolean
    Dim JSON_Index2D As Long
    Dim JSON_LBound2D As Long
    Dim JSON_UBound2D As Long
    Dim JSON_IsFirstItem2D As Boolean
    Dim JSON_Key As Variant
    Dim JSON_Value As Variant
    
    JSON_LBound = -1
    JSON_UBound = -1
    JSON_IsFirstItem = True
    JSON_LBound2D = -1
    JSON_UBound2D = -1
    JSON_IsFirstItem2D = True

    Select Case VBA.VarType(JSON_DictionaryCollectionOrArray)
    Case VBA.vbNull, VBA.vbEmpty
        ConvertToJSON = "null"
    Case VBA.vbDate
        ' TODO Verify date formatting
        ConvertToJSON = """" & VBA.CStr(JSON_DictionaryCollectionOrArray) & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If JSON_ConvertLargeNumbersFromString And JSON_StringIsLargeNumber(JSON_DictionaryCollectionOrArray) Then
            ConvertToJSON = JSON_DictionaryCollectionOrArray
        Else
            ConvertToJSON = """" & JSON_Encode(JSON_DictionaryCollectionOrArray) & """"
        End If
    Case VBA.vbBoolean
        If JSON_DictionaryCollectionOrArray Then
            ConvertToJSON = "true"
        Else
            ConvertToJSON = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        ' Array
        JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength
        
        On Error Resume Next
        
        JSON_LBound = LBound(JSON_DictionaryCollectionOrArray, 1)
        JSON_UBound = UBound(JSON_DictionaryCollectionOrArray, 1)
        JSON_LBound2D = LBound(JSON_DictionaryCollectionOrArray, 2)
        JSON_UBound2D = UBound(JSON_DictionaryCollectionOrArray, 2)
        
        If JSON_LBound >= 0 And JSON_UBound >= 0 Then
            For JSON_Index = JSON_LBound To JSON_UBound
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                If JSON_LBound2D >= 0 And JSON_UBound2D >= 0 Then
                    JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength
                
                    For JSON_Index2D = JSON_LBound2D To JSON_UBound2D
                        If JSON_IsFirstItem2D Then
                            JSON_IsFirstItem2D = False
                        Else
                            JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                        End If
                        
                        JSON_BufferAppend JSON_Buffer, _
                            ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Index, JSON_Index2D), _
                                JSON_ConvertLargeNumbersFromString), _
                            JSON_BufferPosition, JSON_BufferLength
                    Next JSON_Index2D
                    
                    JSON_BufferAppend JSON_Buffer, "]", JSON_BufferPosition, JSON_BufferLength
                    JSON_IsFirstItem2D = True
                Else
                    JSON_BufferAppend JSON_Buffer, _
                        ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Index), _
                            JSON_ConvertLargeNumbersFromString), _
                        JSON_BufferPosition, JSON_BufferLength
                End If
            Next JSON_Index
        End If
        
        On Error GoTo 0
        
        JSON_BufferAppend JSON_Buffer, "]", JSON_BufferPosition, JSON_BufferLength
        
        ConvertToJSON = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
    
    ' Dictionary or Collection
    Case VBA.vbObject
        ' Dictionary
        If VBA.TypeName(JSON_DictionaryCollectionOrArray) = "Dictionary" Then
            JSON_BufferAppend JSON_Buffer, "{", JSON_BufferPosition, JSON_BufferLength
            For Each JSON_Key In JSON_DictionaryCollectionOrArray.Keys
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                JSON_BufferAppend JSON_Buffer, _
                    """" & JSON_Key & """:" & ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Key), JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Key
            JSON_BufferAppend JSON_Buffer, "}", JSON_BufferPosition, JSON_BufferLength
        
        ' Collection
        ElseIf VBA.TypeName(JSON_DictionaryCollectionOrArray) = "Collection" Then
            JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength
            For Each JSON_Value In JSON_DictionaryCollectionOrArray
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                JSON_BufferAppend JSON_Buffer, _
                    ConvertToJSON(JSON_Value, JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Value
            JSON_BufferAppend JSON_Buffer, "]", JSON_BufferPosition, JSON_BufferLength
        End If
        
        ConvertToJSON = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
    Case Else
        ' Number
        On Error Resume Next
        ConvertToJSON = JSON_DictionaryCollectionOrArray
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function JSON_ParseObject(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Dictionary
    Dim JSON_Key As String
    Dim JSON_NextChar As String
    
    Set JSON_ParseObject = New Dictionary
    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '{'")
    Else
        JSON_Index = JSON_Index + 1
        
        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If VBA.Mid$(JSON_String, JSON_Index, 1) = "}" Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 1) = "," Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If
            
            JSON_Key = JSON_ParseKey(JSON_String, JSON_Index)
            JSON_NextChar = JSON_Peek(JSON_String, JSON_Index)
            If JSON_NextChar = "[" Or JSON_NextChar = "{" Then
                Set JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
            Else
                JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
            End If
        Loop
    End If
End Function

Private Function JSON_ParseArray(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Collection
    Set JSON_ParseArray = New Collection
    
    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '['")
    Else
        JSON_Index = JSON_Index + 1
        
        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If VBA.Mid$(JSON_String, JSON_Index, 1) = "]" Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 1) = "," Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If
            
            JSON_ParseArray.Add JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
        Loop
    End If
End Function

Private Function JSON_ParseValue(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Variant
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case VBA.Mid$(JSON_String, JSON_Index, 1)
    Case "{"
        Set JSON_ParseValue = JSON_ParseObject(JSON_String, JSON_Index)
    Case "["
        Set JSON_ParseValue = JSON_ParseArray(JSON_String, JSON_Index)
    Case """", "'"
        JSON_ParseValue = JSON_ParseString(JSON_String, JSON_Index)
    Case Else
        If VBA.Mid$(JSON_String, JSON_Index, 4) = "true" Then
            JSON_ParseValue = True
            JSON_Index = JSON_Index + 4
        ElseIf VBA.Mid$(JSON_String, JSON_Index, 5) = "false" Then
            JSON_ParseValue = False
            JSON_Index = JSON_Index + 5
        ElseIf VBA.Mid$(JSON_String, JSON_Index, 4) = "null" Then
            JSON_ParseValue = Null
            JSON_Index = JSON_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(JSON_String, JSON_Index, 1)) Then
            JSON_ParseValue = JSON_ParseNumber(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
        Else
            Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function JSON_ParseString(JSON_String As String, ByRef JSON_Index As Long) As String
    Dim JSON_Quote As String
    Dim JSON_Char As String
    Dim JSON_Code As String
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    
    JSON_SkipSpaces JSON_String, JSON_Index
    
    ' Store opening quote to look for matching closing quote
    JSON_Quote = VBA.Mid$(JSON_String, JSON_Index, 1)
    JSON_Index = JSON_Index + 1
    
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
        
        Select Case JSON_Char
        Case "\"
            ' Escaped string, \\, or \/
            JSON_Index = JSON_Index + 1
            JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
            
            Select Case JSON_Char
            Case """", "\", "/", "'"
                JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "b"
                JSON_BufferAppend JSON_Buffer, vbBack, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "f"
                JSON_BufferAppend JSON_Buffer, vbFormFeed, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "n"
                JSON_BufferAppend JSON_Buffer, vbCrLf, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "r"
                JSON_BufferAppend JSON_Buffer, vbCr, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "t"
                JSON_BufferAppend JSON_Buffer, vbTab, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                JSON_Index = JSON_Index + 1
                JSON_Code = VBA.Mid$(JSON_String, JSON_Index, 4)
                JSON_BufferAppend JSON_Buffer, VBA.ChrW(VBA.Val("&h" + JSON_Code)), JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 4
            End Select
        Case JSON_Quote
            JSON_ParseString = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
            JSON_Index = JSON_Index + 1
            Exit Function
        Case Else
            JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
            JSON_Index = JSON_Index + 1
        End Select
    Loop
End Function

Private Function JSON_ParseNumber(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Variant
    Dim JSON_Char As String
    Dim JSON_Value As String
    
    JSON_SkipSpaces JSON_String, JSON_Index
    
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
        
        If VBA.InStr("+-0123456789.eE", JSON_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            JSON_Value = JSON_Value & JSON_Char
            JSON_Index = JSON_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15 characters containing only numbers and decimal points -> Number
            If JSON_ConvertLargeNumbersToString And Len(JSON_Value) >= 16 Then
                JSON_ParseNumber = JSON_Value
            Else
                ' Guard for regional settings that use "," for decimal
                ' CStr(0.1) -> "0.1" or "0,1" based on regional settings -> Replace "." with "." or ","
                JSON_Value = VBA.Replace(JSON_Value, ".", VBA.Mid$(VBA.CStr(0.1), 2, 1))
                JSON_ParseNumber = VBA.Val(JSON_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function JSON_ParseKey(JSON_String As String, ByRef JSON_Index As Long) As String
    ' Parse key with single or double quotes
    JSON_ParseKey = JSON_ParseString(JSON_String, JSON_Index)
    
    ' Check for colon and skip if present or throw if not present
    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting ':'")
    Else
        JSON_Index = JSON_Index + 1
    End If
End Function

Private Function JSON_Encode(ByVal JSON_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim JSON_Index As Long
    Dim JSON_Char As String
    Dim JSON_AscCode As Long
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    
    For JSON_Index = 1 To VBA.Len(JSON_Text)
        JSON_Char = VBA.Mid$(JSON_Text, JSON_Index, 1)
        JSON_AscCode = VBA.AscW(JSON_Char)
        
        Select Case JSON_AscCode
        ' " -> 34 -> \"
        Case 34
            JSON_Char = "\"""
        ' \ -> 92 -> \\
        Case 92
            JSON_Char = "\\"
        ' / -> 47 -> \/
        Case 47
            JSON_Char = "\/"
        ' backspace -> 8 -> \b
        Case 8
            JSON_Char = "\b"
        ' form feed -> 12 -> \f
        Case 12
            JSON_Char = "\f"
        ' line feed -> 10 -> \n
        Case 10
            JSON_Char = "\n"
        ' carriage return -> 13 -> \r
        Case 13
            JSON_Char = "\r"
        ' tab -> 9 -> \t
        Case 9
            JSON_Char = "\t"
        ' Non-ascii characters -> convert to 4-digit hex
        Case 0 To 31, 127 To 65535
            JSON_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(JSON_AscCode), 4)
        End Select
            
        JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
    Next JSON_Index
    
    JSON_Encode = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
End Function

Private Function JSON_Peek(JSON_String As String, ByVal JSON_Index As Long, Optional JSON_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing JSON_Index (ByVal instead of ByRef)
    JSON_SkipSpaces JSON_String, JSON_Index
    JSON_Peek = VBA.Mid$(JSON_String, JSON_Index, JSON_NumberOfCharacters)
End Function

Private Sub JSON_SkipSpaces(JSON_String As String, ByRef JSON_Index As Long)
    ' Increment index to skip over spaces
    Do While JSON_Index > 0 And JSON_Index <= VBA.Len(JSON_String) And VBA.Mid$(JSON_String, JSON_Index, 1) = " "
        JSON_Index = JSON_Index + 1
    Loop
End Sub

Private Function JSON_StringIsLargeNumber(JSON_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See JSON_ParseNumber)
    
    Dim JSON_Length As Long
    JSON_Length = VBA.Len(JSON_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If JSON_Length >= 16 And JSON_Length <= 100 Then
        Dim JSON_CharCode As String
        Dim JSON_Index As Long
        
        JSON_StringIsLargeNumber = True
        
        For i = 1 To JSON_Length
            JSON_CharCode = VBA.asc(VBA.Mid$(JSON_String, i, 1))
            Select Case JSON_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                JSON_StringIsLargeNumber = False
                Exit Function
            End Select
        Next i
    End If
End Function

Private Function JSON_ParseErrorMessage(JSON_String As String, ByRef JSON_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['
    
    Dim JSON_StartIndex As Long
    Dim JSON_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    JSON_StartIndex = JSON_Index - 10
    JSON_StopIndex = JSON_Index + 10
    If JSON_StartIndex <= 0 Then
        JSON_StartIndex = 1
    End If
    If JSON_StopIndex > VBA.Len(JSON_String) Then
        JSON_StopIndex = VBA.Len(JSON_String)
    End If

    JSON_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(JSON_String, JSON_StartIndex, JSON_StopIndex - JSON_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(JSON_Index - JSON_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub JSON_BufferAppend(ByRef JSON_Buffer As String, _
                              ByRef JSON_Append As Variant, _
                              ByRef JSON_BufferPosition As Long, _
                              ByRef JSON_BufferLength As Long)
#If Mac Then
    JSON_Buffer = JSON_Buffer & JSON_Append
#Else
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Copy memory for "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

    Dim JSON_AppendLength As Long
    Dim JSON_LengthPlusPosition As Long
    
    JSON_AppendLength = VBA.LenB(JSON_Append)
    JSON_LengthPlusPosition = JSON_AppendLength + JSON_BufferPosition
    
    If JSON_LengthPlusPosition > JSON_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim JSON_TemporaryLength As Long
        
        JSON_TemporaryLength = JSON_BufferLength
        Do While JSON_TemporaryLength < JSON_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If JSON_TemporaryLength = 0 Then
                JSON_TemporaryLength = JSON_TemporaryLength + 510
            Else
                JSON_TemporaryLength = JSON_TemporaryLength + 16384
            End If
        Loop
        
        JSON_Buffer = JSON_Buffer & VBA.Space$((JSON_TemporaryLength - JSON_BufferLength) \ 2)
        JSON_BufferLength = JSON_TemporaryLength
    End If
    
    ' Copy memory from append to buffer at buffer position
    JSON_CopyMemory ByVal JSON_UnsignedAdd(StrPtr(JSON_Buffer), _
                    JSON_BufferPosition), _
                    ByVal StrPtr(JSON_Append), _
                    JSON_AppendLength
    
    JSON_BufferPosition = JSON_BufferPosition + JSON_AppendLength
#End If
End Sub

Private Function JSON_BufferToString(ByRef JSON_Buffer As String, ByVal JSON_BufferPosition As Long, ByVal JSON_BufferLength As Long) As String
#If Mac Then
    JSON_BufferToString = JSON_Buffer
#Else
    If JSON_BufferPosition > 0 Then
        JSON_BufferToString = VBA.Left$(JSON_Buffer, JSON_BufferPosition \ 2)
    End If
#End If
End Function

#If Win64 Then
Private Function JSON_UnsignedAdd(JSON_Start As LongPtr, JSON_Increment As Long) As LongPtr
#Else
Private Function JSON_UnsignedAdd(JSON_Start As Long, JSON_Increment As Long) As Long
#End If

    If JSON_Start And &H80000000 Then
        JSON_UnsignedAdd = JSON_Start + JSON_Increment
    ElseIf (JSON_Start Or &H80000000) < -JSON_Increment Then
        JSON_UnsignedAdd = JSON_Start + JSON_Increment
    Else
        JSON_UnsignedAdd = (JSON_Start + &H80000000) + (JSON_Increment + &H80000000)
    End If
End Function

