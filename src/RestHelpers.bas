Attribute VB_Name = "RestHelpers"
''
' RestHelpers v3.1.4
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
' vba-json
' --------------------------------------------- '

#If Mac Then
Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As Long
Private Declare Function pclose Lib "libc.dylib" (ByVal File As Long) As Long
Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal Items As Long, ByVal stream As Long) As Long
Private Declare Function feof Lib "libc.dylib" (ByVal File As Long) As Long
#End If

Public Type ShellResult
    Output As String
    ExitCode As Long
End Type

Private pDocumentHelper As Object
Private pElHelper As Object
Private pAsyncRequests As Dictionary

' Moved to top from JSONLib
Private Const INVALID_JSON      As Long = 1
Private Const INVALID_OBJECT    As Long = 2
Private Const INVALID_ARRAY     As Long = 3
Private Const INVALID_BOOLEAN   As Long = 4
Private Const INVALID_NULL      As Long = 5
Private Const INVALID_KEY       As Long = 6

Public Enum StatusCodes
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
Public Sub LogRequest(Request As RestRequest)
    If EnableLogging Then
        Debug.Print "--> Request - " & Format(Now, "Long Time")
        Debug.Print Request.MethodName & " " & Request.FullUrl
        
        Dim HeaderKey As Variant
        For Each HeaderKey In Request.Headers.Keys()
            Debug.Print HeaderKey & ": " & Request.Headers(HeaderKey)
        Next HeaderKey
        
        Dim CookieKey As Variant
        For Each CookieKey In Request.Cookies.Keys()
            Debug.Print "Cookie: " & CookieKey & "=" & Request.Cookies(CookieKey)
        Next CookieKey
        
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
        
        Dim Header As Dictionary
        For Each Header In Response.Headers
            Debug.Print Header("key") & ": " & Header("value")
        Next Header
        
        Dim CookieKey As Variant
        For Each CookieKey In Response.Cookies.Keys()
            Debug.Print "Cookie: " & CookieKey & "=" & Response.Cookies(CookieKey)
        Next CookieKey
        
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
Public Function ParseJSON(json As String) As Object
    Set ParseJSON = json_parse(json)
End Function

''
' Convert object to JSON string
'
' @param {Variant} Obj
' @return {String}
' --------------------------------------------- '
Public Function ConvertToJSON(Obj As Variant) As String
    ConvertToJSON = json_toString(Obj)
End Function

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
' Convert dictionary to url encoded string
' TODO: Handle arrays and collections
'
' @param {Variant} Obj
' @return {String} UrlEncoded string (e.g. a=123&b=456&...)
' --------------------------------------------- '
Public Function ConvertToUrlEncoded(Obj As Variant) As String
    If IsArray(Obj) Then
        ' TODO Handle arrays and collections
        Err.Raise vbObjectError + 1, "RestHelpers.ConvertToUrlEncoded", "Arrays are not currently supported by ConvertToUrlEncoded"
    End If
    
    Dim Encoded As String
    Dim ParameterKey As Variant
    Dim Value As Variant
    
    For Each ParameterKey In Obj.Keys()
        If Len(Encoded) > 0 Then: Encoded = Encoded & "&"
        Value = Obj(ParameterKey)
        
        ' Convert boolean to lowercase
        If VarType(Value) = vbBoolean Then
            If Value Then
                Value = "true"
            Else
                Value = "false"
            End If
        End If
        
        Encoded = Encoded & UrlEncode(ParameterKey, True) & "=" & UrlEncode(Value, True)
    Next ParameterKey
    
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
' @param {AvailableFormats} Format
' @return {Object}
' --------------------------------------------- '
Public Function ParseByFormat(Value As String, Format As AvailableFormats) As Object
    Select Case Format
    Case AvailableFormats.json
        Set ParseByFormat = ParseJSON(Value)
    Case AvailableFormats.formurlencoded
        Set ParseByFormat = ParseUrlEncoded(Value)
    Case AvailableFormats.xml
        Set ParseByFormat = ParseXML(Value)
    End Select
End Function

''
' Convert object to given format
'
' @param {Variant} Obj
' @param {AvailableFormats} Format
' @return {String}
' --------------------------------------------- '
Public Function ConvertToFormat(Obj As Variant, Format As AvailableFormats) As String
    Select Case Format
    Case AvailableFormats.json
        ConvertToFormat = ConvertToJSON(Obj)
    Case AvailableFormats.formurlencoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case AvailableFormats.xml
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
            CharCode = Asc(Char)
            
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
    Base64Encode = Replace(BytesToBase64(StringToBytes(Text)), vbLf, "")
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
' Check if protocol is included with url
'
' @param {String} Url
' @return {String} Found protocol
' --------------------------------------------- '
Public Function IncludesProtocol(Url As String) As String
    Dim Parts As New Dictionary
    Set Parts = UrlParts(Url)
    
    If Parts("Protocol") <> "" Then
        IncludesProtocol = Parts("Protocol") & "://"
    End If
End Function

''
' Remove protocol from url (if present)
'
' @param {String} Url
' @return {String} Url without protocol
' --------------------------------------------- '
Public Function RemoveProtocol(Url As String) As String
    Dim Protocol As String
    
    RemoveProtocol = Url
    Protocol = IncludesProtocol(RemoveProtocol)
    If Protocol <> "" Then
        RemoveProtocol = Replace(RemoveProtocol, Protocol, "")
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
' Combine two objects
'
' @param {Dictionary} OriginalObj Original object to add values to
' @param {Dictionary} NewObj New object containing values to add to original object
' @param {Boolean} [OverwriteOriginal=True] Overwrite any values that already exist in the original object
' @return {Dictionary} Combined object
' --------------------------------------------- '
Public Function CombineObjects(ByVal OriginalObj As Dictionary, ByVal NewObj As Dictionary, _
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
    
    Set CombineObjects = Combined
End Function

''
' Apply whitelist to given object to filter out unwanted key/values
'
' @param {Dictionary} Original model to filter
' @param {Variant} WhiteList Array of values to retain in the model
' @return {Dictionary} Filtered object
' --------------------------------------------- '
Public Function FilterObject(ByVal Original As Dictionary, Whitelist As Variant) As Dictionary
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
    
    Set FilterObject = Filtered
End Function

''
' Sort dictionary
' TODO
'
' Source: http://www.cpearson.com/excel/CollectionsAndDictionaries.htm
'         http://www.cpearson.com/excel/SortingArrays.aspx
' --------------------------------------------- '
Public Function SortDictionary(ByVal Dict As Dictionary, SortByKey As Boolean, _
    Optional Descending As Boolean = False, Optional CompareMode As VbCompareMethod = vbTextCompare) As Dictionary
    
    Set SortDictionary = Dict
End Function

''
' Check if given is an array
'
' @param {Object} Obj
' @return {Boolean}
' --------------------------------------------- '
Public Function IsArray(Obj As Variant) As Boolean
    If Not IsEmpty(Obj) Then
        If IsObject(Obj) Then
            If TypeOf Obj Is Collection Then
                IsArray = True
            End If
        ElseIf VarType(Obj) = vbArray Or VarType(Obj) = 8204 Then
            ' VarType = 8204 seems to arise from Array(...) constructor
            IsArray = True
        End If
    End If
End Function

''
' Add or update key/value in dictionary
'
' @param {Dictionary} Dict
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddToDictionary(ByRef Dict As Dictionary, Key As String, Value As Variant)
    If Not Dict.Exists(Key) Then
        Dict.Add Key, Value
    Else
        Dict(Key) = Value
    End If
End Sub

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
    Dim HeaderKey As Variant
    For Each HeaderKey In Request.Headers.Keys()
        Http.setRequestHeader HeaderKey, Request.Headers(HeaderKey)
    Next HeaderKey
    
    Dim CookieKey As Variant
    For Each CookieKey In Request.Cookies.Keys()
        Http.setRequestHeader "Cookie", CookieKey & "=" & Request.Cookies(CookieKey)
    Next CookieKey
End Sub

''
' Create simple response
'
' @param {StatusCodes} StatusCode
' @param {String} StatusDescription
' @return {RestResponse}
' --------------------------------------------- '
Public Function CreateResponse(StatusCode As StatusCodes, StatusDescription As String) As RestResponse
    Set CreateResponse = New RestResponse
    CreateResponse.StatusCode = StatusCode
    CreateResponse.StatusDescription = StatusDescription
End Function

''
' Create response for http
'
' @param {WinHttpRequest} Http
' @param {AvailableFormats} [Format=json]
' @return {RestResponse}
' --------------------------------------------- '
Public Function CreateResponseFromHttp(ByRef Http As Object, Optional Format As AvailableFormats = AvailableFormats.json) As RestResponse
    Set CreateResponseFromHttp = New RestResponse
    
    CreateResponseFromHttp.StatusCode = Http.Status
    CreateResponseFromHttp.StatusDescription = Http.StatusText
    CreateResponseFromHttp.Body = Http.ResponseBody
    CreateResponseFromHttp.Content = Http.ResponseText
    
    ' Convert content to data by format
    If Format <> AvailableFormats.plaintext Then
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
Public Function CreateResponseFromCURL(Result As ShellResult, Optional Format As AvailableFormats = AvailableFormats.json) As RestResponse
    Dim StatusCode As Long
    Dim StatusText As String
    Dim Headers As String
    Dim Body As Variant
    Dim ResponseText As String
    
    If Result.ExitCode > 0 Then
        Dim ErrorNumber As Long
        
        ErrorNumber = Result.ExitCode / 256
        ' 7 - CURLE_COULDNT_CONNECT
        ' 28 - CURLE_OPERATION_TIMEDOUT
        If ErrorNumber = 7 Or ErrorNumber = 28 Then
            Set CreateResponseFromCURL = CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
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
    Body = StringToBytes(ResponseText)
    
    ' Create Response
    Set CreateResponseFromCURL = New RestResponse
    CreateResponseFromCURL.StatusCode = StatusCode
    CreateResponseFromCURL.StatusDescription = StatusText
    CreateResponseFromCURL.Body = Body
    CreateResponseFromCURL.Content = ResponseText
    
    ' Convert content to data by format
    If Format <> AvailableFormats.plaintext Then
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
                Set Header = New Dictionary
                Header.Add "key", Trim(Mid$(Lines(i), 1, InStr(1, Lines(i), ":") - 1))
                Header.Add "value", Trim(Mid$(Lines(i), InStr(1, Lines(i), ":") + 1, Len(Lines(i))))
            Else
                Header("value") = Header("value") & vbCrLf & Lines(i)
            End If
        End If
    Next i
    
    Set ExtractHeaders = Headers
End Function

''
' Extract cookies from response headers
'
' @param {Collection} Headers
' @return {Dictionary} Cookies
' --------------------------------------------- '
Public Function ExtractCookies(Headers As Collection) As Dictionary
    Dim Cookies As New Dictionary
    Dim Cookie As String
    Dim Key As String
    Dim Value As String
    Dim Header As Dictionary
    
    For Each Header In Headers
        If Header("key") = "Set-Cookie" Then
            Cookie = Header("value")
            Key = Mid$(Cookie, 1, InStr(1, Cookie, "=") - 1)
            Value = Mid$(Cookie, InStr(1, Cookie, "=") + 1, Len(Cookie))
            
            If InStr(1, Value, ";") Then
                Value = Mid$(Value, 1, InStr(1, Value, ";") - 1)
            End If
            
            If Cookies.Exists(Key) Then
                Cookies(Key) = UrlDecode(Value)
            Else
                Cookies.Add Key, UrlDecode(Value)
            End If
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
' Get name for format
'
' @param {AvailableFormats} Format
' @return {String}
' --------------------------------------------- '
Public Function FormatToName(Format As AvailableFormats) As String
    Select Case Format
    Case AvailableFormats.formurlencoded
        FormatToName = "form-urlencoded"
    Case AvailableFormats.json
        FormatToName = "json"
    Case AvailableFormats.xml
        FormatToName = "xml"
    Case AvailableFormats.plaintext
        FormatToName = "txt"
    End Select
End Function

''
' Get content-type for format
'
' @param {AvailableFormats} Format
' @return {String}
' --------------------------------------------- '
Public Function FormatToContentType(Format As AvailableFormats) As String
    Select Case Format
    Case AvailableFormats.formurlencoded
        FormatToContentType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case AvailableFormats.json
        FormatToContentType = "application/json"
    Case AvailableFormats.xml
        FormatToContentType = "application/xml"
    Case AvailableFormats.plaintext
        FormatToContentType = "text/plain"
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
        
        LogDebug "Async Timeout: " & AsyncWrapper.Request.FullUrl, "RestHelpers.TimeoutTimerExpired"
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
    HMACSHA1 = BytesToFormat(HMACSHA1AsBytes(Text, Secret), Format)
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
    HMACSHA256 = BytesToFormat(HMACSHA256AsBytes(Text, Secret), Format)
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
    MD5 = BytesToFormat(MD5AsBytes(Text), Format)
End Function


Public Function HMACSHA1AsBytes(Text As String, Secret As String) As Byte()
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -sha1 -hmac " & PrepareTextForShell(Secret)
    HMACSHA1AsBytes = HexToBytes(ExecuteInShell(Command).Output)
#Else
    Dim Crypto As Object
    Set Crypto = CreateObject("System.Security.Cryptography.HMACSHA1")
    
    Crypto.Key = StringToBytes(Secret)
    HMACSHA1AsBytes = Crypto.ComputeHash_2(StringToBytes(Text))
#End If
End Function

Public Function HMACSHA256AsBytes(Text As String, Secret As String) As Byte()
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -sha256 -hmac " & PrepareTextForShell(Secret)
    HMACSHA256AsBytes = HexToBytes(ExecuteInShell(Command).Output)
#Else
    Dim Crypto As Object
    Set Crypto = CreateObject("System.Security.Cryptography.HMACSHA256")
    
    Crypto.Key = StringToBytes(Secret)
    HMACSHA256AsBytes = Crypto.ComputeHash_2(StringToBytes(Text))
#End If
End Function

Public Function MD5AsBytes(Text As String) As Byte()
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -md5"
    MD5AsBytes = HexToBytes(ExecuteInShell(Command).Output)
#Else
    Dim Crypto As Object
    Set Crypto = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    MD5AsBytes = Crypto.ComputeHash_2(StringToBytes(Text))
#End If
End Function

''
' Convert string to bytes
'
' @param {String} Text
' @return {Byte()}
' --------------------------------------------- '
Public Function StringToBytes(Text As String) As Byte()
    StringToBytes = StrConv(Text, vbFromUnicode)
End Function

Public Function BytesToString(Bytes() As Byte) As String
    BytesToString = StrConv(Bytes, vbUnicode)
End Function

Public Function BytesToHex(Bytes() As Byte) As String
    Dim i As Integer
    For i = LBound(Bytes) To UBound(Bytes)
        BytesToHex = BytesToHex & LCase(Right("0" & Hex$(Bytes(i)), 2))
    Next i
End Function

Public Function HexToBytes(Hex As String) As Byte()
    Dim Bytes() As Byte
    Dim HexIndex As Integer
    Dim ByteIndex As Integer
    
    ' Remove linefeeds
    Hex = Replace(Hex, vbLf, "")
    
    ReDim Bytes(Len(Hex) / 2 - 1)
    ByteIndex = 0
    For HexIndex = 1 To Len(Hex) Step 2
        Bytes(ByteIndex) = CLng("&H" & Mid$(Hex, HexIndex, 2))
        ByteIndex = ByteIndex + 1
    Next HexIndex
    
    HexToBytes = Bytes
End Function

Public Function BytesToBase64(Bytes() As Byte) As String
#If Mac Then
    Dim Command As String
    Command = "printf " & PrepareTextForShell(BytesToString(Bytes)) & " | openssl base64"
    BytesToBase64 = ExecuteInShell(Command).Output
#Else
    Dim xml As Object
    Dim Node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")

    ' byte array to base64
    Set Node = xml.createElement("b64")
    Node.DataType = "bin.base64"
    Node.nodeTypedValue = Bytes
    BytesToBase64 = Node.Text

    Set Node = Nothing
    Set xml = Nothing
#End If
End Function

''
' Convert bytes to given format (Hex or Base64)
'
' @param {Byte()} Bytes
' @param {String} Format (Hex or Base64)
' @return {String}
' --------------------------------------------- '
Public Function BytesToFormat(Bytes() As Byte, Format As String) As String
    Select Case UCase(Format)
    Case "HEX"
        BytesToFormat = BytesToHex(Bytes)
    Case "BASE64"
        BytesToFormat = BytesToBase64(Bytes)
    End Select
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

' ======================================================================================== '
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
' Changes for Excel-REST:
' - Updated json_parseNumber to reduce chance of overflow
' - Swapped Mid for Mid$
' - Handle colon in object key
' - Handle duplicate keys in object parsing
' - Change methods to Private and prefix with json_
'
' ======================================================================================== '

' (Moved to top of file)
'Private Const INVALID_JSON      As Long = 1
'Private Const INVALID_OBJECT    As Long = 2
'Private Const INVALID_ARRAY     As Long = 3
'Private Const INVALID_BOOLEAN   As Long = 4
'Private Const INVALID_NULL      As Long = 5
'Private Const INVALID_KEY       As Long = 6

'
'   parse string and create JSON object (Dictionary or Collection in VB)
'
Private Function json_parse(ByRef Str As String) As Object

    Dim Index As Long
    Index = 1
    
    On Error Resume Next

    Call json_skipChar(Str, Index)
    Select Case Mid$(Str, Index, 1)
    Case "{"
        Set json_parse = json_parseObject(Str, Index)
    Case "["
        Set json_parse = json_parseArray(Str, Index)
    End Select

End Function

'
'   parse collection of key/value (Dictionary in VB)
'
Private Function json_parseObject(ByRef Str As String, ByRef Index As Long) As Dictionary

    Set json_parseObject = New Dictionary
    
    ' "{"
    Call json_skipChar(Str, Index)
    If Mid$(Str, Index, 1) <> "{" Then Err.Raise vbObjectError + INVALID_OBJECT, Description:="char " & Index & " : " & Mid$(Str, Index)
    Index = Index + 1
    
    Dim Key As String
    
    Do
        Call json_skipChar(Str, Index)
        If "}" = Mid$(Str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = Mid$(Str, Index, 1) Then
            Index = Index + 1
            Call json_skipChar(Str, Index)
        End If
        
        Key = json_parseKey(Str, Index)
        If Not json_parseObject.Exists(Key) Then
            json_parseObject.Add Key, json_parseValue(Str, Index)
        Else
            json_parseObject.Item(Key) = json_parseValue(Str, Index)
        End If
    Loop

End Function

'
'   parse list (Collection in VB)
'
Private Function json_parseArray(ByRef Str As String, ByRef Index As Long) As Collection

    Set json_parseArray = New Collection
    
    ' "["
    Call json_skipChar(Str, Index)
    If Mid$(Str, Index, 1) <> "[" Then Err.Raise vbObjectError + INVALID_ARRAY, Description:="char " & Index & " : " + Mid$(Str, Index)
    Index = Index + 1
    
    Do
        
        Call json_skipChar(Str, Index)
        If "]" = Mid$(Str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = Mid$(Str, Index, 1) Then
            Index = Index + 1
            Call json_skipChar(Str, Index)
        End If
        
        ' add value
        json_parseArray.Add json_parseValue(Str, Index)
        
    Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function json_parseValue(ByRef Str As String, ByRef Index As Long)

    Call json_skipChar(Str, Index)
    
    Select Case Mid$(Str, Index, 1)
    Case "{"
        Set json_parseValue = json_parseObject(Str, Index)
    Case "["
        Set json_parseValue = json_parseArray(Str, Index)
    Case """", "'"
        json_parseValue = json_parseString(Str, Index)
    Case "t", "f"
        json_parseValue = json_parseBoolean(Str, Index)
    Case "n"
        json_parseValue = json_parseNull(Str, Index)
    Case Else
        json_parseValue = json_parseNumber(Str, Index)
    End Select

End Function

'
'   parse string
'
Private Function json_parseString(ByRef Str As String, ByRef Index As Long) As String

    Dim quote   As String
    Dim Char    As String
    Dim Code    As String
    
    Call json_skipChar(Str, Index)
    quote = Mid$(Str, Index, 1)
    Index = Index + 1
    Do While Index > 0 And Index <= Len(Str)
        Char = Mid$(Str, Index, 1)
        Select Case (Char)
        Case "\"
            Index = Index + 1
            Char = Mid$(Str, Index, 1)
            Select Case (Char)
            Case """", "\", "/" ' Before: Case """", "\\", "/"
                json_parseString = json_parseString & Char
                Index = Index + 1
            Case "b"
                json_parseString = json_parseString & vbBack
                Index = Index + 1
            Case "f"
                json_parseString = json_parseString & vbFormFeed
                Index = Index + 1
            Case "n"
                json_parseString = json_parseString & vbNewLine
                Index = Index + 1
            Case "r"
                json_parseString = json_parseString & vbCr
                Index = Index + 1
            Case "t"
                json_parseString = json_parseString & vbTab
                Index = Index + 1
            Case "u"
                Index = Index + 1
                Code = Mid$(Str, Index, 4)
                json_parseString = json_parseString & ChrW(Val("&h" + Code))
                Index = Index + 4
            End Select
        Case quote
            
            Index = Index + 1
            Exit Function
        Case Else
            json_parseString = json_parseString & Char
            Index = Index + 1
        End Select
    Loop

End Function

'
'   parse number
'
Private Function json_parseNumber(ByRef Str As String, ByRef Index As Long)

    Dim Value   As String
    Dim Char    As String
    
    Call json_skipChar(Str, Index)
    Do While Index > 0 And Index <= Len(Str)
        Char = Mid$(Str, Index, 1)
        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            Index = Index + 1
        Else
            json_parseNumber = Val(Value)
            Exit Function
        End If
    Loop


End Function

'
'   parse true / false
'
Private Function json_parseBoolean(ByRef Str As String, ByRef Index As Long) As Boolean

    Call json_skipChar(Str, Index)
    If Mid$(Str, Index, 4) = "true" Then
        json_parseBoolean = True
        Index = Index + 4
    ElseIf Mid$(Str, Index, 5) = "false" Then
        json_parseBoolean = False
        Index = Index + 5
    Else
        Err.Raise vbObjectError + INVALID_BOOLEAN, Description:="char " & Index & " : " & Mid$(Str, Index)
    End If

End Function

'
'   parse null
'
Private Function json_parseNull(ByRef Str As String, ByRef Index As Long)

    Call json_skipChar(Str, Index)
    If Mid$(Str, Index, 4) = "null" Then
        json_parseNull = Null
        Index = Index + 4
    Else
        Err.Raise vbObjectError + INVALID_NULL, Description:="char " & Index & " : " & Mid$(Str, Index)
    End If

End Function

Private Function json_parseKey(ByRef Str As String, ByRef Index As Long) As String

    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim Char    As String
    
    Call json_skipChar(Str, Index)
    Do While Index > 0 And Index <= Len(Str)
        Char = Mid$(Str, Index, 1)
        Select Case (Char)
        Case """"
            dquote = Not dquote
            Index = Index + 1
            If Not dquote Then
                Call json_skipChar(Str, Index)
                If Mid$(Str, Index, 1) <> ":" Then
                    Err.Raise vbObjectError + INVALID_KEY, Description:="char " & Index & " : " & json_parseKey
                End If
            End If
        Case "'"
            squote = Not squote
            Index = Index + 1
            If Not squote Then
                Call json_skipChar(Str, Index)
                If Mid$(Str, Index, 1) <> ":" Then
                    Err.Raise vbObjectError + INVALID_KEY, Description:="char " & Index & " : " & json_parseKey
                End If
            End If
        Case ":"
            If Not dquote And Not squote Then
                Index = Index + 1
                Exit Do
            Else
                ' Colon in key name
                json_parseKey = json_parseKey & Char
                Index = Index + 1
            End If
        Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
            Else
                json_parseKey = json_parseKey & Char
            End If
            Index = Index + 1
        End Select
    Loop

End Function

'
'   skip special character
'
Private Sub json_skipChar(ByRef Str As String, ByRef Index As Long)

    While Index > 0 And Index <= Len(Str) And InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Mid$(Str, Index, 1))
        Index = Index + 1
    Wend

End Sub

Private Function json_toString(ByRef Obj As Variant) As String

    Select Case VarType(Obj)
        Case vbNull
            json_toString = "null"
        Case vbEmpty
            json_toString = "null"
        Case vbDate
            json_toString = """" & CStr(Obj) & """"
        Case vbString
            json_toString = """" & json_encode(Obj) & """"
        Case vbObject
            Dim bFI, i
            bFI = True
            If TypeName(Obj) = "Dictionary" Then
                json_toString = json_toString & "{"
                Dim Keys
                Keys = Obj.Keys
                For i = 0 To Obj.Count - 1
                    If bFI Then bFI = False Else json_toString = json_toString & ","
                    Dim Key
                    Key = Keys(i)
                    json_toString = json_toString & """" & Key & """:" & json_toString(Obj(Key))
                Next i
                json_toString = json_toString & "}"
            ElseIf TypeName(Obj) = "Collection" Then
                json_toString = json_toString & "["
                Dim Value
                For Each Value In Obj
                    If bFI Then bFI = False Else json_toString = json_toString & ","
                    json_toString = json_toString & json_toString(Value)
                Next Value
                json_toString = json_toString & "]"
            End If
        Case vbBoolean
            If Obj Then json_toString = "true" Else json_toString = "false"
        Case vbVariant, vbArray, vbArray + vbVariant
            Dim sEB
            json_toString = json_multiArray(Obj, 1, "", sEB)
        Case Else
            json_toString = Replace(Obj, ",", ".")
    End Select

End Function

Private Function json_encode(Str) As String
    
    Dim i, j, aL1, aL2, C, p

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To Len(Str)
        p = True
        C = Mid$(Str, i, 1)
        For j = 0 To 7
            If C = Chr(aL1(j)) Then
                json_encode = json_encode & "\" & Chr(aL2(j))
                p = False
                Exit For
            End If
        Next

        If p Then
            Dim A
            A = AscW(C)
            If A > 31 And A < 127 Then
                json_encode = json_encode & C
            ElseIf A > -1 Or A < 65535 Then
                json_encode = json_encode & "\u" & String(4 - Len(Hex(A)), "0") & Hex(A)
            End If
        End If
    Next
End Function

Private Function json_multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
    Dim iDU, iDL, i ' Integer DimensionUBound, Integer DimensionLBound
    On Error Resume Next
    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)
    
    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
    If Err.Number = 9 Then
        sPB1 = sPT & sPS
        For i = 1 To Len(sPB1)
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid$(sPB1, i, 1)
        Next
'        json_multiArray = json_multiArray & json_toString(Eval("aBD(" & sPB2 & ")"))
        json_multiArray = json_multiArray & json_toString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        json_multiArray = json_multiArray & "["
        For i = iDL To iDU
            json_multiArray = json_multiArray & json_multiArray(aBD, iBC + 1, i, sPT)
            If i < iDU Then json_multiArray = json_multiArray & ","
        Next
        json_multiArray = json_multiArray & "]"
        sPT = Left(sPT, iBC - 2)
    End If
    Err.Clear
End Function

