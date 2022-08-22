Attribute VB_Name = "WebHelpers"
''
' WebHelpers v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Contains general-purpose helpers that are used throughout VBA-Web. Includes:
'
' - Logging
' - Converters and encoding
' - Url handling
' - Object/Dictionary/Collection/Array helpers
' - Request preparation / handling
' - Timing
' - Mac
' - Cryptography
' - Converters (JSON, XML, Url-Encoded)
'
' Errors:
' 11000 - Error during parsing
' 11001 - Error during conversion
' 11002 - No matching converter has been registered
' 11003 - Error while getting url parts
' 11099 - XML format is not currently supported
'
' @module WebHelpers
' @author tim.hall.engr@gmail.com
' @author: Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' VBA-Git Annotations
' https://github.com/VBA-Tools-v2/VBA-Git | https://radiuscore.co.nz
'
' @excludeobfuscation
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder VBA-Web.Helpers
'@ignoremodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

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
' VBA-JSON
' VBA-UTC
' AutoProxy
' --------------------------------------------- '

' Custom formatting uses the standard version of Application.Run,
' which is incompatible with some Office applications (e.g. Word 2011 for Mac)
'
' If you have compilation errors in ParseByFormat or ConvertToFormat,
' you can disable custom formatting by setting the following compiler flag to False
#Const EnableCustomFormatting = True

' === AutoProxy Headers
#If Mac Then
#ElseIf VBA7 Then

Private Declare PtrSafe Sub AutoProxy_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal AutoProxy_lpDest As LongPtr, ByVal AutoProxy_lpSource As LongPtr, ByVal AutoProxy_cbCopy As Long)
Private Declare PtrSafe Function AutoProxy_SysAllocString Lib "oleaut32" Alias "SysAllocString" _
    (ByVal AutoProxy_pwsz As LongPtr) As LongPtr
Private Declare PtrSafe Function AutoProxy_GlobalFree Lib "kernel32" Alias "GlobalFree" _
    (ByVal AutoProxy_p As LongPtr) As LongPtr
Private Declare PtrSafe Function AutoProxy_GetIEProxy Lib "WinHTTP.dll" Alias "WinHttpGetIEProxyConfigForCurrentUser" _
    (ByRef AutoProxy_proxyConfig As AUTOPROXY_IE_PROXY_CONFIG) As Long
Private Declare PtrSafe Function AutoProxy_GetProxyForUrl Lib "WinHTTP.dll" Alias "WinHttpGetProxyForUrl" _
    (ByVal AutoProxy_hSession As LongPtr, ByVal AutoProxy_pszUrl As LongPtr, ByRef AutoProxy_pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef AutoProxy_pProxyInfo As AUTOPROXY_INFO) As Long
Private Declare PtrSafe Function AutoProxy_HttpOpen Lib "WinHTTP.dll" Alias "WinHttpOpen" _
    (ByVal AutoProxy_pszUserAgent As LongPtr, ByVal AutoProxy_dwAccessType As Long, ByVal AutoProxy_pszProxyName As LongPtr, ByVal AutoProxy_pszProxyBypass As LongPtr, ByVal AutoProxy_dwFlags As Long) As LongPtr
Private Declare PtrSafe Function AutoProxy_HttpClose Lib "WinHTTP.dll" Alias "WinHttpCloseHandle" _
    (ByVal AutoProxy_hInternet As LongPtr) As Long

Private Type AUTOPROXY_IE_PROXY_CONFIG
    AutoProxy_fAutoDetect As Long
    AutoProxy_lpszAutoConfigUrl As LongPtr
    AutoProxy_lpszProxy As LongPtr
    AutoProxy_lpszProxyBypass As LongPtr
End Type
Private Type AUTOPROXY_OPTIONS
    AutoProxy_dwFlags As Long
    AutoProxy_dwAutoDetectFlags As Long
    AutoProxy_lpszAutoConfigUrl As LongPtr
    AutoProxy_lpvReserved As LongPtr
    AutoProxy_dwReserved As Long
    AutoProxy_fAutoLogonIfChallenged As Long
End Type
Private Type AUTOPROXY_INFO
    AutoProxy_dwAccessType As Long
    AutoProxy_lpszProxy As LongPtr
    AutoProxy_lpszProxyBypass As LongPtr
End Type

#Else

Private Declare Sub AutoProxy_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal AutoProxy_lpDest As Long, ByVal AutoProxy_lpSource As Long, ByVal AutoProxy_cbCopy As Long)
Private Declare Function AutoProxy_SysAllocString Lib "oleaut32" Alias "SysAllocString" _
    (ByVal AutoProxy_pwsz As Long) As Long
Private Declare Function AutoProxy_GlobalFree Lib "kernel32" Alias "GlobalFree" _
    (ByVal AutoProxy_p As Long) As Long
Private Declare Function AutoProxy_GetIEProxy Lib "WinHTTP.dll" Alias "WinHttpGetIEProxyConfigForCurrentUser" _
    (ByRef AutoProxy_proxyConfig As AUTOPROXY_IE_PROXY_CONFIG) As Long
Private Declare Function AutoProxy_GetProxyForUrl Lib "WinHTTP.dll" Alias "WinHttpGetProxyForUrl" _
    (ByVal AutoProxy_hSession As Long, ByVal AutoProxy_pszUrl As Long, ByRef AutoProxy_pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef AutoProxy_pProxyInfo As AUTOPROXY_INFO) As Long
Private Declare Function AutoProxy_HttpOpen Lib "WinHTTP.dll" Alias "WinHttpOpen" _
    (ByVal AutoProxy_pszUserAgent As Long, ByVal AutoProxy_dwAccessType As Long, ByVal AutoProxy_pszProxyName As Long, ByVal AutoProxy_pszProxyBypass As Long, ByVal AutoProxy_dwFlags As Long) As Long
Private Declare Function AutoProxy_HttpClose Lib "WinHTTP.dll" Alias "WinHttpCloseHandle" _
    (ByVal AutoProxy_hInternet As Long) As Long

Private Type AUTOPROXY_IE_PROXY_CONFIG
    AutoProxy_fAutoDetect As Long
    AutoProxy_lpszAutoConfigUrl As Long
    AutoProxy_lpszProxy As Long
    AutoProxy_lpszProxyBypass As Long
End Type
Private Type AUTOPROXY_OPTIONS
    AutoProxy_dwFlags As Long
    AutoProxy_dwAutoDetectFlags As Long
    AutoProxy_lpszAutoConfigUrl As Long
    AutoProxy_lpvReserved As Long
    AutoProxy_dwReserved As Long
    AutoProxy_fAutoLogonIfChallenged As Long
End Type
Private Type AUTOPROXY_INFO
    AutoProxy_dwAccessType As Long
    AutoProxy_lpszProxy As Long
    AutoProxy_lpszProxyBypass As Long
End Type

#End If

#If Mac Then
#Else
' Constants for dwFlags of AUTOPROXY_OPTIONS
Const AUTOPROXY_AUTO_DETECT = 1
Const AUTOPROXY_CONFIG_URL = 2

' Constants for dwAutoDetectFlags
Const AUTOPROXY_DETECT_TYPE_DHCP = 1
Const AUTOPROXY_DETECT_TYPE_DNS = 2
#End If
' === End AutoProxy

#If Mac Then
#If VBA7 Then
Private Declare PtrSafe Function web_popen Lib "/usr/lib/libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As LongPtr
Private Declare PtrSafe Function web_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" (ByVal web_File As LongPtr) As LongPtr
Private Declare PtrSafe Function web_fread Lib "/usr/lib/libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As LongPtr, ByVal web_Items As LongPtr, ByVal web_Stream As LongPtr) As LongPtr
Private Declare PtrSafe Function web_feof Lib "/usr/lib/libc.dylib" Alias "feof" (ByVal web_File As LongPtr) As LongPtr
#Else
Private Declare Function web_popen Lib "libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As Long
Private Declare Function web_pclose Lib "libc.dylib" Alias "pclose" (ByVal web_File As Long) As Long
Private Declare Function web_fread Lib "libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As Long, ByVal web_Items As Long, ByVal web_Stream As Long) As Long
Private Declare Function web_feof Lib "libc.dylib" Alias "feof" (ByVal web_File As Long) As Long
#End If
#End If

Public Const WebUserAgent As String = "VBA-Web v4.1.6 (https://github.com/VBA-tools/VBA-Web)"

' @internal
Public Type ShellResult
    Output As String
    ExitCode As Long
End Type

Private web_pDocumentHelper As Object
Private web_pElHelper As Object
Private web_pConverters As Dictionary

' --------------------------------------------- '
' Types and Properties
' --------------------------------------------- '

''
' Helper for common http status codes. (Use underlying status code for any codes not listed)
'
' @example
' ```VB.net
' Dim Response As WebResponse
'
' If Response.StatusCode = WebStatusCode.Ok Then
'   ' Ok
' ElseIf Response.StatusCode = 418 Then
'   ' I'm a teapot
' End If
' ```
'
' @property WebStatusCode
' @param Ok `200`
' @param Created `201`
' @param NoContent `204`
' @param NotModified `304`
' @param BadRequest `400`
' @param Unauthorized `401`
' @param Forbidden `403`
' @param NotFound `404`
' @param RequestTimeout `408`
' @param UnsupportedMediaType `415`
' @param InternalServerError `500`
' @param BadGateway `502`
' @param ServiceUnavailable `503`
' @param GatewayTimeout `504`
''
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

''
' @property WebMethod
' @param HttpGet
' @param HttpPost
' @param HttpGet
' @param HttpGet
' @param HttpGet
' @default HttpGet
''
Public Enum WebMethod
    HttpGet = 0
    HttpPost = 1
    HttpPut = 2
    HttpDelete = 3
    HttpPatch = 4
    HttpHead = 5
End Enum

''
' @property WebFormat
' @param PlainText
' @param Json
' @param FormUrlEncoded
' @param Xml
' @param Custom
' @default PlainText
''
Public Enum WebFormat
    PlainText = 0
    Json = 1
    FormUrlEncoded = 2
    Xml = 3
    Custom = 9
End Enum

''
' @property UrlEncodingMode
' @param StrictUrlEncoding RFC 3986, ALPHA / DIGIT / "-" / "." / "_" / "~"
' @param FormUrlEncoding ALPHA / DIGIT / "-" / "." / "_" / "*", (space) -> "+", &...; UTF-8 encoding
' @param QueryUrlEncoding Subset of strict and form that should be suitable for non-form-urlencoded query strings
'   ALPHA / DIGIT / "-" / "." / "_"
' @param CookieUrlEncoding strict / "!" / "#" / "$" / "&" / "'" / "(" / ")" / "*" / "+" /
'   "/" / ":" / "<" / "=" / ">" / "?" / "@" / "[" / "]" / "^" / "`" / "{" / "|" / "}"
' @param PathUrlEncoding strict / "!" / "$" / "&" / "'" / "(" / ")" / "*" / "+" / "," / ";" / "=" / ":" / "@"
''
Public Enum UrlEncodingMode
    StrictUrlEncoding
    FormUrlEncoding
    QueryUrlEncoding
    CookieUrlEncoding
    PathUrlEncoding
End Enum

''
' Enable logging of requests and responses and other internal messages from VBA-Web.
' Should be the first step in debugging VBA-Web if something isn't working as expected.
' (Logs display in Immediate Window (`View > Immediate Window` or `ctrl+g`)
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://api.example.com/v1/"
'
' Dim RequestWithTypo As New WebRequest
' RequestWithTypo.Resource = "peeple/{id}"
' RequestWithType.AddUrlSegment "idd", 123
'
' ' Enable logging before the request is executed
' WebHelpers.EnableLogging = True
'
' Dim Response As WebResponse
' Set Response = Client.Execute(Request)
'
' ' Immediate window:
' ' --> Request - (Time)
' ' GET https://api.example.com/v1/peeple/{id}
' ' Headers...
' '
' ' <-- Response - (Time)
' ' 404 ...
' ```
'
' @property EnableLogging
' @type Boolean
' @default False
''
Public EnableLogging As Boolean

''
' Store currently running async requests
'
' @property AsyncRequests
' @type Dictionary
''
Public AsyncRequests As Dictionary

' ============================================= '
' 1. Logging
' ============================================= '

''
' Log message (when logging is enabled with `EnableLogging`)
' with optional location where the message is coming from.
' Useful when writing extensions to VBA-Web (like an `IWebAuthenticator`).
'
' @example
' ```VB.net
' LogDebug "Executing request..."
' ' -> VBA-Web: Executing request...
'
' LogDebug "Executing request...", "Module.Function"
' ' -> Module.Function: Executing request...
' ```
'
' @method LogDebug
' @param {String} Message
' @param {String} [From="VBA-Web"]
''
Public Sub LogDebug(Message As String, Optional From As String = "VBA-Web")
    If EnableLogging Then
        Debug.Print From & ": " & Message
    End If
End Sub

''
' Log warning (even when logging is disabled with `EnableLogging`)
' with optional location where the message is coming from.
' Useful when writing extensions to VBA-Web (like an `IWebAuthenticator`).
'
' @example
' ```VB.net
' WebHelpers.LogWarning "Something could go wrong"
' ' -> WARNING - VBA-Web: Something could go wrong
'
' WebHelpers.LogWarning "Something could go wrong", "Module.Function"
' ' -> WARNING - Module.Function: Something could go wrong
' ```
'
' @method LogWarning
' @param {String} Message
' @param {String} [From="VBA-Web"]
''
Public Sub LogWarning(Message As String, Optional From As String = "VBA-Web")
    Debug.Print "WARNING - " & From & ": " & Message
End Sub

''
' Log error (even when logging is disabled with `EnableLogging`)
' with optional location where the message is coming from and error number.
' Useful when writing extensions to VBA-Web (like an `IWebAuthenticator`).
'
' @example
' ```VB.net
' WebHelpers.LogError "Something went wrong"
' ' -> ERROR - VBA-Web: Something went wrong
'
' WebHelpers.LogError "Something went wrong", "Module.Function"
' ' -> ERROR - Module.Function: Something went wrong
'
' WebHelpers.LogError "Something went wrong", "Module.Function", 100
' ' -> ERROR - Module.Function: 100, Something went wrong
' ```
'
' @method LogError
' @param {String} Message
' @param {String} [From="VBA-Web"]
' @param {Long} [ErrNumber=0]
''
Public Sub LogError(Message As String, Optional From As String = "VBA-Web", Optional ErrNumber As Long = 0)
    Dim web_ErrorValue As String
    If ErrNumber <> 0 Then
        web_ErrorValue = ErrNumber

        If ErrNumber < 0 Then
            web_ErrorValue = web_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If

        web_ErrorValue = web_ErrorValue & ", "
    End If

    Debug.Print "ERROR - " & From & ": " & web_ErrorValue & Message
End Sub

''
' Log details of the request (Url, headers, cookies, body, etc.).
'
' @method LogRequest
' @param {WebClient} Client
' @param {WebRequest} Request
''
Public Sub LogRequest(Client As WebClient, Request As WebRequest)
    If EnableLogging Then
        Debug.Print "--> Request - " & Format(Now, "Long Time")
        Debug.Print MethodToName(Request.Method) & " " & Client.GetFullUrl(Request)

        Dim web_KeyValue As Dictionary
        For Each web_KeyValue In Request.Headers
            Debug.Print web_KeyValue("Key") & ": " & web_KeyValue("Value")
        Next web_KeyValue

        For Each web_KeyValue In Request.Cookies
            Debug.Print "Cookie: " & web_KeyValue("Key") & "=" & web_KeyValue("Value")
        Next web_KeyValue

        If Not IsEmpty(Request.Body) Then
            Debug.Print vbNewLine & CStr(Request.Body)
        End If

        Debug.Print
    End If
End Sub

''
' Log details of the response (Status, headers, content, etc.).
'
' @method LogResponse
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {WebResponse} Response
''
Public Sub LogResponse(Client As WebClient, Request As WebRequest, Response As WebResponse)
    If EnableLogging Then
        Dim web_KeyValue As Dictionary

        Debug.Print "<-- Response - " & Format(Now, "Long Time")
        Debug.Print Response.StatusCode & " " & Response.StatusDescription

        For Each web_KeyValue In Response.Headers
            Debug.Print web_KeyValue("Key") & ": " & web_KeyValue("Value")
        Next web_KeyValue

        For Each web_KeyValue In Response.Cookies
            Debug.Print "Cookie: " & web_KeyValue("Key") & "=" & web_KeyValue("Value")
        Next web_KeyValue

        Debug.Print vbNewLine & Response.Content & vbNewLine
    End If
End Sub

''
' Obfuscate any secure information before logging.
'
' @example
' ```VB.net
' Dim Password As String
' Password = "Secret"
'
' WebHelpers.LogDebug "Password = " & WebHelpers.Obfuscate(Password)
' -> Password = ******
' ```
'
' @param {String} Secure Message to obfuscate
' @param {String} [Character = *] Character to obfuscate with
' @return {String}
''
Public Function Obfuscate(Secure As String, Optional Character As String = "*") As String
    Obfuscate = VBA.String$(VBA.Len(Secure), Character)
End Function

' ============================================= '
' 2. Converters and encoding
' ============================================= '

'
' Parse JSON value to `Dictionary` if it's an object or `Collection` if it's an array.
'
' @method ParseJson
' @param {String} Json JSON value to parse
' @return {Dictionary|Collection}
'
' (Implemented in VBA-JSON)

'
' Convert `Dictionary`, `Collection`, or `Array` to JSON string.
'
' @method ConvertToJson
' @param {Dictionary|Collection|Array} Obj
' @return {String}
'
' (Implemented in VBA-JSON)

''
' Parse Url-Encoded value to `Dictionary`.
'
' @method ParseUrlEncoded
' @param {String} UrlEncoded Url-Encoded value to parse
' @return {Dictionary} Parsed
''
Public Function ParseUrlEncoded(Encoded As String) As Dictionary
    Dim web_Items As Variant
    Dim web_i As Integer
    Dim web_Parts As Variant
    Dim web_Key As String
    Dim web_Value As Variant
    Dim web_Parsed As New Dictionary

    web_Items = VBA.Split(Encoded, "&")
    For web_i = LBound(web_Items) To UBound(web_Items)
        web_Parts = VBA.Split(web_Items(web_i), "=")

        If UBound(web_Parts) - LBound(web_Parts) >= 1 Then
            ' TODO: Handle numbers, arrays, and object better here
            web_Key = UrlDecode(VBA.CStr(web_Parts(LBound(web_Parts))))
            web_Value = UrlDecode(VBA.CStr(web_Parts(LBound(web_Parts) + 1)))

            web_Parsed(web_Key) = web_Value
        End If
    Next web_i

    Set ParseUrlEncoded = web_Parsed
End Function

''
' Convert `Dictionary`/`Collection` to Url-Encoded string.
'
' @method ConvertToUrlEncoded
' @param {Dictionary|Collection|Variant} Obj Value to convert to Url-Encoded string
' @return {String} UrlEncoded string (e.g. a=123&b=456&...)
''
Public Function ConvertToUrlEncoded(Obj As Variant, Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.FormUrlEncoding) As String
    Dim web_Encoded As String

    If TypeOf Obj Is Collection Then
        Dim web_KeyValue As Dictionary

        For Each web_KeyValue In Obj
            If VBA.Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_KeyValue("Key"), web_KeyValue("Value"), EncodingMode)
        Next web_KeyValue
    Else
        Dim web_Key As Variant

        For Each web_Key In Obj.Keys()
            If Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_Key, Obj(web_Key), EncodingMode)
        Next web_Key
    End If

    ConvertToUrlEncoded = web_Encoded
End Function

''
' Parse XML string to `Dictionary` or `DOMDocument`.
'
' @method ParseXml
' @param {String} Encoded XML value to parse
' @return {Dictionary|DOMDocument}
'
' (Implemented in VBA-XML)
''

''
' Convert object (Dictionary/Collection/DOMDocument) to XML string.
'
' @method ConvertToXml
' @param {Dictionary|Collection|DOMDocument} XmlValue
' @param {Integer|String} Whitespace "Pretty" print XML with given number of spaces per indentation (Integer) or given string
' @return {String}
'
' (Implemented in VBA-XML)
''

''
' Helper for parsing value to given `WebFormat` or custom format.
' Returns `Dictionary` or `Collection` based on given `Value`.
'
' @method ParseByFormat
' @param {String} Value Value to parse
' @param {WebFormat} Format
' @param {String} [CustomFormat=""] Name of registered custom converter
' @param {Variant} [Bytes] Bytes for custom convert (if `ParseType = "Binary"`)
' @return {Dictionary|Collection|Object}
' @throws 11000 - Error during parsing
''
Public Function ParseByFormat(Value As String, Format As WebFormat, _
    Optional CustomFormat As String = "", Optional Bytes As Variant) As Object

    On Error GoTo web_ErrorHandling

    ' Don't attempt to parse blank values
    If Value = "" And CustomFormat = "" Then
        Exit Function
    End If

    Select Case Format
    Case WebFormat.Json
        Set ParseByFormat = ParseJson(Value)
    Case WebFormat.FormUrlEncoded
        Set ParseByFormat = ParseUrlEncoded(Value)
    Case WebFormat.Xml
        Set ParseByFormat = ParseXml(Value)
    Case WebFormat.Custom
#If EnableCustomFormatting Then
        Dim web_Converter As Dictionary
        Dim web_Callback As String

        Set web_Converter = web_GetConverter(CustomFormat)
        web_Callback = web_Converter("ParseCallback")

        If web_Converter.Exists("Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter("Instance")

            If web_Converter("ParseType") = "Binary" Then
                Set ParseByFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Bytes)
            Else
                Set ParseByFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Value)
            End If
        Else
            If web_Converter("ParseType") = "Binary" Then
                Set ParseByFormat = Application.Run(web_Callback, Bytes)
            Else
                Set ParseByFormat = Application.Run(web_Callback, Value)
            End If
        End If
#Else
    LogWarning "Custom formatting is disabled. To use WebFormat.Custom, enable custom formatting with the EnableCustomFormatting flag in WebHelpers"
#End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during parsing" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.ParseByFormat", 11000
    Err.Raise 11000, "WebHelpers.ParseByFormat", web_ErrorDescription
End Function

''
' Helper for converting value to given `WebFormat` or custom format.
'
' _Note_ Only some converters handle `Collection` or `Array`.
'
' @method ConvertToFormat
' @param {Dictionary|Collection|Variant} Obj
' @param {WebFormat} Format
' @param {String} [CustomFormat] Name of registered custom converter
' @return {Variant}
' @throws 11001 - Error during conversion
''
Public Function ConvertToFormat(Obj As Variant, Format As WebFormat, Optional CustomFormat As String = "") As Variant
    On Error GoTo web_ErrorHandling

    Select Case Format
    Case WebFormat.Json
        ConvertToFormat = ConvertToJson(Obj)
    Case WebFormat.FormUrlEncoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case WebFormat.Xml
        ConvertToFormat = ConvertToXml(Obj)
    Case WebFormat.Custom
#If EnableCustomFormatting Then
        Dim web_Converter As Dictionary
        Dim web_Callback As String

        Set web_Converter = web_GetConverter(CustomFormat)
        web_Callback = web_Converter("ConvertCallback")

        If web_Converter.Exists("Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter("Instance")
            ConvertToFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Obj)
        Else
            ConvertToFormat = Application.Run(web_Callback, Obj)
        End If
#Else
    LogWarning "Custom formatting is disabled. To use WebFormat.Custom, enable custom formatting with the EnableCustomFormatting flag in WebHelpers"
#End If
    Case Else
        If VBA.VarType(Obj) = vbString Then
            ' Plain text
            ConvertToFormat = Obj
        End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during conversion" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.ConvertToFormat", 11001
    Err.Raise 11001, "WebHelpers.ConvertToFormat", web_ErrorDescription
End Function

''
' Encode string for URLs
'
' See https://github.com/VBA-tools/VBA-Web/wiki/Url-Encoding for details
'
' References:
' - RFC 3986, https://tools.ietf.org/html/rfc3986
' - form-urlencoded encoding algorithm,
'   https://www.w3.org/TR/html5/forms.html#application/x-www-form-urlencoded-encoding-algorithm
' - RFC 6265 (Cookies), https://tools.ietf.org/html/rfc6265
'   Note: "%" is allowed in spec, but is currently excluded due to parsing issues
'
' @method UrlEncode
' @param {Variant} Text Text to encode
' @param {Boolean} [SpaceAsPlus = False] `%20` if `False` / `+` if `True`
'   DEPRECATED Use EncodingMode:=FormUrlEncoding
' @param {Boolean} [EncodeUnsafe = True] Encode characters that could be misunderstood within URLs.
'   (``SPACE, ", <, >, #, %, {, }, |, \, ^, ~, `, [, ]``)
'   DEPRECATED This was based on an outdated URI spec and has since been removed.
'     EncodingMode:=CookieUrlEncoding is the closest approximation of this behavior
' @param {UrlEncodingMode} [EncodingMode = StrictUrlEncoding]
' @return {String} Encoded string
''
Public Function UrlEncode(Text As Variant, _
    Optional SpaceAsPlus As Boolean = False, Optional EncodeUnsafe As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    If SpaceAsPlus = True Then
        LogWarning "SpaceAsPlus is deprecated and will be removed in VBA-Web v5. " & _
            "Use EncodingMode:=FormUrlEncoding instead", "WebHelpers.UrlEncode"
    End If
    If EncodeUnsafe = False Then
        LogWarning "EncodeUnsafe has been removed as it was based on an outdated url encoding specification. " & _
            "Use EncodingMode:=CookieUrlEncoding to approximate this behavior", "WebHelpers.UrlEncode"
    End If

    Dim web_UrlVal As String
    Dim web_StringLen As Long

    web_UrlVal = VBA.CStr(Text)
    web_StringLen = VBA.Len(web_UrlVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Integer
        Dim web_Char As String
        Dim web_Space As String
        ReDim web_Result(web_StringLen)

        ' StrictUrlEncoding - ALPHA / DIGIT / "-" / "." / "_" / "~"
        ' FormUrlEncoding   - ALPHA / DIGIT / "-" / "." / "_" / "*" / (space) -> "+"
        ' QueryUrlEncoding  - ALPHA / DIGIT / "-" / "." / "_"
        ' CookieUrlEncoding - strict / "!" / "#" / "$" / "&" / "'" / "(" / ")" / "*" / "+" /
        '   "/" / ":" / "<" / "=" / ">" / "?" / "@" / "[" / "]" / "^" / "`" / "{" / "|" / "}"
        ' PathUrlEncoding   - strict / "!" / "$" / "&" / "'" / "(" / ")" / "*" / "+" / "," / ";" / "=" / ":" / "@"

        ' Set space value
        If SpaceAsPlus Or EncodingMode = UrlEncodingMode.FormUrlEncoding Then
            web_Space = "+"
        Else
            web_Space = "%20"
        End If

        ' Loop through string characters
        For web_i = 1 To web_StringLen
            ' Get character and ascii code
            web_Char = VBA.Mid$(web_UrlVal, web_i, 1)
            web_CharCode = VBA.Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    ' ALPHA
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    ' DIGIT
                    web_Result(web_i) = web_Char
                Case 45, 46, 95
                    ' "-" / "." / "_"
                    web_Result(web_i) = web_Char

                Case 32
                    ' (space)
                    ' FormUrlEncoding -> "+"
                    ' Else -> "%20"
                    web_Result(web_i) = web_Space

                Case 33, 36, 38, 39, 40, 41, 43, 58, 61, 64
                    ' "!" / "$" / "&" / "'" / "(" / ")" / "+" / ":" / "=" / "@"
                    ' PathUrlEncoding, CookieUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 35, 45, 46, 47, 60, 62, 63, 91, 93, 94, 95, 96, 123, 124, 125
                    ' "#" / "-" / "." / "/" / "<" / ">" / "?" / "[" / "]" / "^" / "_" / "`" / "{" / "|" / "}"
                    ' CookieUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 42
                    ' "*"
                    ' FormUrlEncoding, PathUrlEncoding, CookieUrlEncoding -> "*"
                    ' Else -> "%2A"
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.PathUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then

                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 44, 59
                    ' "," / ";"
                    ' PathUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 126
                    ' "~"
                    ' FormUrlEncoding, QueryUrlEncoding -> "%7E"
                    ' Else -> "~"
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding Or EncodingMode = UrlEncodingMode.QueryUrlEncoding Then
                        web_Result(web_i) = "%7E"
                    Else
                        web_Result(web_i) = web_Char
                    End If

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)

                ' TODO For non-ASCII characters,
                '
                ' FormUrlEncoded:
                '
                ' Replace the character by a string consisting of a U+0026 AMPERSAND character (&), a "#" (U+0023) character,
                ' one or more ASCII digits representing the Unicode code point of the character in base ten, and finally a ";" (U+003B) character.
                '
                ' Else:
                '
                ' Encode to sequence of 2 or 3 bytes in UTF-8, then percent encode
                ' Reference Implementation: https://www.w3.org/International/URLUTF8Encoder.java
            End Select
        Next web_i
        UrlEncode = VBA.Join$(web_Result, "")
    End If
End Function

''
' Decode Url-encoded string.
'
' @method UrlDecode
' @param {String} Encoded Text to decode
' @param {Boolean} [PlusAsSpace = True] Decode plus as space
'   DEPRECATED Use EncodingMode:=FormUrlEncoding Or QueryUrlEncoding
' @param {UrlEncodingMode} [EncodingMode = StrictUrlEncoding]
' @return {String} Decoded string
''
Public Function UrlDecode(Encoded As String, _
    Optional PlusAsSpace As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    Dim web_StringLen As Long
    web_StringLen = VBA.Len(Encoded)

    If web_StringLen > 0 Then
        Dim web_i As Long
        Dim web_Result As String
        Dim web_Temp As String

        For web_i = 1 To web_StringLen
            web_Temp = VBA.Mid$(Encoded, web_i, 1)

            If web_Temp = "+" And _
                (PlusAsSpace _
                 Or EncodingMode = UrlEncodingMode.FormUrlEncoding _
                 Or EncodingMode = UrlEncodingMode.QueryUrlEncoding) Then

                web_Temp = " "
            ElseIf web_Temp = "%" And web_StringLen >= web_i + 2 Then
                web_Temp = VBA.Mid$(Encoded, web_i + 1, 2)
                web_Temp = VBA.Chr(VBA.CInt("&H" & web_Temp))

                web_i = web_i + 2
            End If

            ' TODO Handle non-ASCII characters

            web_Result = web_Result & web_Temp
        Next web_i

        UrlDecode = web_Result
    End If
End Function

''
' Base64-encode text.
'
' @param {Variant} Text Text to encode
' @return {String} Encoded string
''
Public Function Base64Encode(Text As String) As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl base64"
    Base64Encode = ExecuteInShell(web_Command).Output
#Else
    Dim web_Bytes() As Byte

    web_Bytes = VBA.StrConv(Text, vbFromUnicode)
    Base64Encode = web_AnsiBytesToBase64(web_Bytes)
#End If

    Base64Encode = VBA.Replace$(Base64Encode, vbLf, "")
End Function

''
' Decode Base64-encoded text
'
' @param {Variant} Encoded Text to decode
' @return {String} Decoded string
''
Public Function Base64Decode(Encoded As Variant) As String
    ' Add trailing padding, if necessary
    If (VBA.Len(Encoded) Mod 4 > 0) Then
        Encoded = Encoded & VBA.Left("====", 4 - (VBA.Len(Encoded) Mod 4))
    End If

#If Mac Then
    Dim web_Command As String
    web_Command = "echo " & PrepareTextForShell(Encoded) & " | openssl base64 -d"
    Base64Decode = ExecuteInShell(web_Command).Output
#Else
    Dim web_XmlObj As Object
    Dim web_Node As Object

    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    Set web_Node = web_XmlObj.createElement("b64")

    web_Node.DataType = "bin.base64"
    web_Node.Text = Encoded
    Base64Decode = VBA.StrConv(web_Node.nodeTypedValue, vbUnicode)

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
#End If
End Function

''
' Register custom converter for converting request `Body` and response `Content`.
' If the `ConvertCallback` or `ParseCallback` are object methods,
' pass in an object instance.
' If the `ParseCallback` needs the raw binary response value (e.g. file download),
' set `ParseType = "Binary"`, otherwise `"String"` is used.
'
' - `ConvertCallback` signature: `Function ...(Value As Variant) As String`
' - `ParseCallback` signature: `Function ...(Value As String) As Object`
'
' @example
' ```VB.net
' ' 1. Use global module functions for Convert and Parse
' ' ---
' ' Module: CSVConverter
' Function ParseCSV(Value As String) As Object
'   ' ...
' End Function
' Function ConvertToCSV(Value As Variant) As String
'   ' ...
' End Function
'
' WebHelpers.RegisterConverter "csv", "text/csv", _
'   "CSVConverter.ConvertToCSV", "CSVConverter.ParseCSV"
'
' ' 2. Use object instance functions for Convert and Parse
' ' ---
' ' Object: CSVConverterClass
' ' same as above...
'
' Dim Converter As New CSVConverterClass
' WebHelpers.RegisterConverter "csv", "text/csv", _
'   "ConvertToCSV", "ParseCSV", Instance:=Converter
'
' ' 3. Pass raw binary value to ParseCallback
' ' ---
' ' Module: ImageConverter
' Function ParseImage(Bytes As Variant) As Object
'   ' ...
' End Function
' Function ConvertToImage(Value As Variant) As String
'   ' ...
' End Function
'
' WebHelpers.RegisterConverter "image", "image/jpeg", _
'   "ImageConverter.ConvertToImage", "ImageConverter.ParseImage", _
'   ParseType:="Binary"
' ```
'
' @method RegisterConverter
' @param {String} Name
'   Name of converter for use with `CustomRequestFormat` or `CustomResponseFormat`
' @param {String} MediaType
'   Media type to use for `Content-Type` and `Accept` headers
' @param {String} ConvertCallback Global or object function name for converting
' @param {String} ParseCallback Global or object function name for parsing
' @param {Object} [Instance]
'   Use instance methods for `ConvertCallback` and `ParseCallback`
' @param {String} [ParseType="String"]
'   "String"` (default) or `"Binary"` to pass raw binary response to `ParseCallback`
''
Public Sub RegisterConverter( _
    Name As String, MediaType As String, ConvertCallback As String, ParseCallback As String, _
    Optional Instance As Object, Optional ParseType As String = "String")

    Dim web_Converter As New Dictionary
    web_Converter("MediaType") = MediaType
    web_Converter("ConvertCallback") = ConvertCallback
    web_Converter("ParseCallback") = ParseCallback
    web_Converter("ParseType") = ParseType

    If Not Instance Is Nothing Then
        Set web_Converter("Instance") = Instance
    End If

    If web_pConverters Is Nothing Then: Set web_pConverters = New Dictionary
    Set web_pConverters(Name) = web_Converter
End Sub

' Helper for getting custom converter
' @throws 11002 - No matching converter has been registered
Private Function web_GetConverter(web_CustomFormat As String) As Dictionary
    If web_pConverters.Exists(web_CustomFormat) Then
        Set web_GetConverter = web_pConverters(web_CustomFormat)
    Else
        LogError "No matching converter has been registered for custom format: " & web_CustomFormat, _
            "WebHelpers.web_GetConverter", 11002
        Err.Raise 11002, "WebHelpers.web_GetConverter", _
            "No matching converter has been registered for custom format: " & web_CustomFormat
    End If
End Function

' ============================================= '
' 3. Url handling
' ============================================= '

''
' Join Url with /
'
' @example
' ```VB.net
' Debug.Print WebHelpers.JoinUrl("a/", "/b")
' Debug.Print WebHelpers.JoinUrl("a", "b")
' Debug.Print WebHelpers.JoinUrl("a/", "b")
' Debug.Print WebHelpers.JoinUrl("a", "/b")
' -> a/b
' ```
'
' @param {String} LeftSide
' @param {String} RightSide
' @return {String} Joined url
''
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
' Get relevant parts of the given url.
' Returns `Protocol`, `Host`, `Port`, `Path`, `Querystring`, and `Hash`
'
' @example
' ```VB.net
' WebHelpers.GetUrlParts "https://www.google.com/a/b/c.html?a=1&b=2#hash"
' ' -> Protocol = "https"
' '    Host = "www.google.com"
' '    Port = "443"
' '    Path = "/a/b/c.html"
' '    Querystring = "a=1&b=2"
' '    Hash = "hash"
'
' WebHelpers.GetUrlParts "localhost:3000/a/b/c"
' ' -> Protocol = ""
' '    Host = "localhost"
' '    Port = "3000"
' '    Path = "/a/b/c"
' '    Querystring = ""
' '    Hash = ""
' ```
'
' @method GetUrlParts
' @param {String} Url
' @return {Dictionary} Parts of url
'   Protocol, Host, Port, Path, Querystring, Hash
' @throws 11003 - Error while getting url parts
''
Public Function GetUrlParts(Url As String) As Dictionary
    Dim web_Parts As New Dictionary

    On Error GoTo web_ErrorHandling

#If Mac Then
    ' Run perl script to parse url

    Dim web_AddedProtocol As Boolean
    Dim web_Command As String
    Dim web_Results As Variant
    Dim web_ResultPart As Variant
    Dim web_EqualsIndex As Long
    Dim web_Key As String
    Dim web_Value As String

    ' Add Protocol if missing
    If InStr(1, Url, "://") <= 0 Then
        web_AddedProtocol = True
        If InStr(1, Url, "//") = 1 Then
            Url = "http" & Url
        Else
            Url = "http://" & Url
        End If
    End If

    web_Command = "perl -e '{use URI::URL;" & vbNewLine & _
        "$url = new URI::URL """ & Url & """;" & vbNewLine & _
        "print ""Protocol="" . $url->scheme;" & vbNewLine & _
        "print "" | Host="" . $url->host;" & vbNewLine & _
        "print "" | Port="" . $url->port;" & vbNewLine & _
        "print "" | FullPath="" . $url->full_path;" & vbNewLine & _
        "print "" | Hash="" . $url->frag;" & vbNewLine & _
    "}'"

    web_Results = Split(ExecuteInShell(web_Command).Output, " | ")
    For Each web_ResultPart In web_Results
        web_EqualsIndex = InStr(1, web_ResultPart, "=")
        web_Key = Trim(VBA.Mid$(web_ResultPart, 1, web_EqualsIndex - 1))
        web_Value = Trim(VBA.Mid$(web_ResultPart, web_EqualsIndex + 1))

        If web_Key = "FullPath" Then
            ' For properly escaped path and querystring, need to use full_path
            ' But, need to split FullPath into Path...?Querystring
            Dim QueryIndex As Integer

            QueryIndex = InStr(1, web_Value, "?")
            If QueryIndex > 0 Then
                web_Parts.Add "Path", Mid$(web_Value, 1, QueryIndex - 1)
                web_Parts.Add "Querystring", Mid$(web_Value, QueryIndex + 1)
            Else
                web_Parts.Add "Path", web_Value
                web_Parts.Add "Querystring", ""
            End If
        Else
            web_Parts.Add web_Key, web_Value
        End If
    Next web_ResultPart

    If web_AddedProtocol And web_Parts.Exists("Protocol") Then
        web_Parts("Protocol") = ""
    End If
#Else
    ' Create document/element is expensive, cache after creation
    If web_pDocumentHelper Is Nothing Or web_pElHelper Is Nothing Then
        Set web_pDocumentHelper = CreateObject("htmlfile")
        Set web_pElHelper = web_pDocumentHelper.createElement("a")
    End If

    web_pElHelper.href = Url
    web_Parts.Add "Protocol", Replace(web_pElHelper.Protocol, ":", "", Count:=1)
    web_Parts.Add "Host", web_pElHelper.hostname
    web_Parts.Add "Port", web_pElHelper.port
    web_Parts.Add "Path", web_pElHelper.pathname
    web_Parts.Add "Querystring", Replace(web_pElHelper.Search, "?", "", Count:=1)
    web_Parts.Add "Hash", Replace(web_pElHelper.Hash, "#", "", Count:=1)
#End If

    If web_Parts("Protocol") = "localhost" Then
        ' localhost:port/... was passed in without protocol
        Dim PathParts As Variant
        PathParts = Split(web_Parts("Path"), "/")

        web_Parts("Port") = PathParts(0)
        web_Parts("Protocol") = ""
        web_Parts("Host") = "localhost"
        web_Parts("Path") = Replace(web_Parts("Path"), web_Parts("Port"), "", Count:=1)
    End If
    If Left(web_Parts("Path"), 1) <> "/" Then
        web_Parts("Path") = "/" & web_Parts("Path")
    End If

    Set GetUrlParts = web_Parts
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while getting url parts" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.GetUrlParts", 11003
    Err.Raise 11003, "WebHelpers.GetUrlParts", web_ErrorDescription
End Function

''
' Converts url formatted querystring (i.e. as retrived from GetUrlParts) to dictionary.
'
' @method QuerystringToDictionary
' @param {String} Querystring
' @retun {Dictionary}
''
Public Function QuerystringToDictionary(ByVal QueryString As String) As Dictionary
    Dim web_Parts As New Dictionary
    Dim web_Queries As Variant
    Dim web_QueryParts As Variant
    
    web_Queries = VBA.Split(QueryString, "&")
    
    For Each web_QueryParts In web_Queries
        web_Parts.Add VBA.Split(web_QueryParts, "=")(0), VBA.Split(web_QueryParts, "=")(1)
    Next web_QueryParts

    Set QuerystringToDictionary = web_Parts
    
    Set web_Parts = Nothing
    Set web_Queries = Nothing
    Set web_QueryParts = Nothing
End Function

' ============================================= '
' 4. Object/Dictionary/Collection/Array helpers
' ============================================= '

''
' Create a cloned copy of the `Dictionary`.
' This is not a deep copy, so children objects are copied by reference.
'
' @method CloneDictionary
' @param {Dictionary} Original
' @return {Dictionary} Clone
''
Public Function CloneDictionary(Original As Dictionary) As Dictionary
    Dim web_Key As Variant

    Set CloneDictionary = New Dictionary
    For Each web_Key In Original.Keys
        CloneDictionary.Add VBA.CStr(web_Key), Original(web_Key)
    Next web_Key
End Function

''
' Create a cloned copy of the `Collection`.
' This is not a deep copy, so children objects are copied by reference.
'
' _Note_ Keys are not transferred to clone
'
' @method CloneCollection
' @param {Collection} Original
' @return {Collection} Clone
''
Public Function CloneCollection(Original As Collection) As Collection
    Dim web_Item As Variant

    Set CloneCollection = New Collection
    For Each web_Item In Original
        CloneCollection.Add web_Item
    Next web_Item
End Function

''
' Helper for creating `Key-Value` pair with `Dictionary`.
' Used in `WebRequest`/`WebResponse` `Cookies`, `Headers`, and `QuerystringParams`
'
' @example
' ```VB.net
' WebHelpers.CreateKeyValue "abc", 123
' ' -> {"Key": "abc", "Value": 123}
' ```
'
' @method CreateKeyValue
' @param {String} Key
' @param {Variant} Value
' @return {Dictionary}
''
Public Function CreateKeyValue(Key As String, Value As Variant) As Dictionary
    Dim web_KeyValue As New Dictionary

    web_KeyValue("Key") = Key
    web_KeyValue("Value") = Value
    Set CreateKeyValue = web_KeyValue
End Function

''
' Search a `Collection` of `KeyValue` and retrieve the value for the given key.
'
' @example
' ```VB.net
' Dim KeyValues As New Collection
' KeyValues.Add WebHelpers.CreateKeyValue("abc", 123)
'
' WebHelpers.FindInKeyValues KeyValues, "abc"
' ' -> 123
'
' WebHelpers.FindInKeyValues KeyValues, "unknown"
' ' -> Empty
' ```
'
' @method FindInKeyValues
' @param {Collection} KeyValues
' @param {Variant} Key to find
' @return {Variant}
''
Public Function FindInKeyValues(KeyValues As Collection, Key As Variant) As Variant
    Dim web_KeyValue As Dictionary

    For Each web_KeyValue In KeyValues
        If web_KeyValue("Key") = Key Then
            FindInKeyValues = web_KeyValue("Value")
            Exit Function
        End If
    Next web_KeyValue
End Function

''
' Helper for adding/replacing `KeyValue` in `Collection` of `KeyValue`
' - Add if key not found
' - Replace if key is found
'
' @example
' ```VB.net
' Dim KeyValues As New Collection
' KeyValues.Add WebHelpers.CreateKeyValue("a", 123)
' KeyValues.Add WebHelpers.CreateKeyValue("b", 456)
' KeyValues.Add WebHelpers.CreateKeyValue("c", 789)
'
' WebHelpers.AddOrReplaceInKeyValues KeyValues, "b", "abc"
' WebHelpers.AddOrReplaceInKeyValues KeyValues, "d", "def"
'
' ' -> [
' '      {"Key":"a","Value":123},
' '      {"Key":"b","Value":"abc"},
' '      {"Key":"c","Value":789},
' '      {"Key":"d","Value":"def"}
' '    ]
' ```
'
' @method AddOrReplaceInKeyValues
' @param {Collection} KeyValues
' @param {Variant} Key
' @param {Variant} Value
' @return {Variant}
''
Public Sub AddOrReplaceInKeyValues(KeyValues As Collection, Key As Variant, Value As Variant)
    Dim web_KeyValue As Dictionary
    Dim web_Index As Long
    Dim web_NewKeyValue As Dictionary

    Set web_NewKeyValue = CreateKeyValue(CStr(Key), Value)

    web_Index = 1
    For Each web_KeyValue In KeyValues
        If web_KeyValue("Key") = Key Then
            ' Replace existing
            KeyValues.Remove web_Index

            If KeyValues.Count = 0 Then
                KeyValues.Add web_NewKeyValue
            ElseIf web_Index > KeyValues.Count Then
                KeyValues.Add web_NewKeyValue, After:=web_Index - 1
            Else
                KeyValues.Add web_NewKeyValue, Before:=web_Index
            End If
            Exit Sub
        End If

        web_Index = web_Index + 1
    Next web_KeyValue

    ' Add
    KeyValues.Add web_NewKeyValue
End Sub

''
' Method to join collection.
'
' @method JoinCollection
' @param {Collection} SourceCollection
' @param {Variant} Delimiter
' @return {String}
''
Public Function JoinCollection(ByVal SourceCollection As Collection, Optional ByVal Delimiter As Variant = ",") As String
    Dim web_Item As Variant
    
    For Each web_Item In SourceCollection
        JoinCollection = JoinCollection & Delimiter & web_Item
    Next web_Item
    
    If VBA.Len(JoinCollection) >= VBA.Len(Delimiter) And Not Delimiter = vbNullString Then
        JoinCollection = VBA.Right$(JoinCollection, VBA.Len(JoinCollection) - VBA.Len(Delimiter))
    End If
End Function

''
' Convert string to a collection.
'
' @method StringToCollection
' @param {String} SourceString
' @param {Variant} Delimiter
' @return {Collection}
''
Public Function StringToCollection(ByVal SourceString As String, Optional ByVal Delimiter As Variant = ",") As Collection
    Dim web_Item As Variant
    Dim web_Collection As Collection
    Set web_Collection = New Collection

    For Each web_Item In VBA.Split(SourceString, Delimiter)
        web_Collection.Add web_Item
    Next web_Item
    
    Set StringToCollection = web_Collection
End Function

''
' Check if given Key exists in Collection.
'
' @param {Collection} TargetCollection
' @param {Variant} Key
' @return {Boolean}
''
Public Function ExistsInCollection(ByVal TargetCollection As Collection, ByVal Key As Variant) As Boolean
    Dim web_Item As Variant
    
    On Error Resume Next
    web_Item = VBA.IsObject(TargetCollection.Item(Key))
    ExistsInCollection = Not VBA.IsEmpty(web_Item)
    
End Function

' ============================================= '
' 5. Request preparation / handling
' ============================================= '

''
' Get the media-type for the given format / custom format.
'
' @method FormatToMediaType
' @param {WebFormat} Format
' @param {String} [CustomFormat] Needed if `Format = WebFormat.Custom`
' @return {String}
''
Public Function FormatToMediaType(Format As WebFormat, Optional CustomFormat As String) As String
    Select Case Format
    Case WebFormat.FormUrlEncoded
        FormatToMediaType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case WebFormat.Json
        FormatToMediaType = "application/json"
    Case WebFormat.Xml
        FormatToMediaType = "application/xml"
    Case WebFormat.Custom
        FormatToMediaType = web_GetConverter(CustomFormat)("MediaType")
    Case Else
        FormatToMediaType = "text/plain"
    End Select
End Function

''
' Get the method name for the given `WebMethod`
'
' @example
' ```VB.net
' WebHelpers.MethodToName WebMethod.HttpPost
' ' -> "POST"
' ```
'
' @method MethodToName
' @param {WebMethod} Method
' @return {String}
''
Public Function MethodToName(Method As WebMethod) As String
    Select Case Method
    Case WebMethod.HttpDelete
        MethodToName = "DELETE"
    Case WebMethod.HttpPut
        MethodToName = "PUT"
    Case WebMethod.HttpPatch
        MethodToName = "PATCH"
    Case WebMethod.HttpPost
        MethodToName = "POST"
    Case WebMethod.HttpGet
        MethodToName = "GET"
    Case WebMethod.HttpHead
        MethodToName = "HEAD"
    End Select
End Function

' ============================================= '
' 6. Timing
' ============================================= '

''
' Handle timeout timers expiring
'
' @internal
' @method OnTimeoutTimerExpired
' @param {String} RequestId
''
Public Sub OnTimeoutTimerExpired(web_RequestId As String)
    If Not AsyncRequests Is Nothing Then
        If AsyncRequests.Exists(web_RequestId) Then
            Dim web_AsyncWrapper As Object
            Set web_AsyncWrapper = AsyncRequests(web_RequestId)
            web_AsyncWrapper.TimedOut
        End If
    End If
End Sub

' ============================================= '
' 7. Mac
' ============================================= '

''
' Execute the given command
'
' @internal
' @method ExecuteInShell
' @param {String} Command
' @return {ShellResult}
''
Public Function ExecuteInShell(web_Command As String) As ShellResult
#If Mac Then
#If VBA7 Then
    Dim web_File As LongPtr
#Else
    Dim web_File As Long
#End If

    Dim web_Chunk As String
    Dim web_Read As Long

    On Error GoTo web_Cleanup

    web_File = web_popen(web_Command, "r")

    If web_File = 0 Then
        ' TODO Investigate why this could happen and what should be done if it happens
        Exit Function
    End If

    Do While web_feof(web_File) = 0
        web_Chunk = VBA.Space$(50)
        web_Read = CLng(web_fread(web_Chunk, 1, Len(web_Chunk) - 1, web_File))
        If web_Read > 0 Then
            web_Chunk = VBA.Left$(web_Chunk, web_Read)
            ExecuteInShell.Output = ExecuteInShell.Output & web_Chunk
        End If
        
        VBA.DoEvents
    Loop

web_Cleanup:

    ExecuteInShell.ExitCode = CLng(web_pclose(web_File))
#End If
End Function

''
' Prepare text for shell
' - Wrap in "..."
' - Replace ! with '!' (reserved in bash)
' - Escape \, `, $, %, and "
'
' @internal
' @method PrepareTextForShell
' @param {String} Text
' @return {String}
''
Public Function PrepareTextForShell(ByVal web_Text As String) As String
    ' Escape special characters (except for !)
    web_Text = VBA.Replace(web_Text, "\", "\\")
    web_Text = VBA.Replace(web_Text, "`", "\`")
    web_Text = VBA.Replace(web_Text, "$", "\$")
    web_Text = VBA.Replace(web_Text, "%", "\%")
    web_Text = VBA.Replace(web_Text, """", "\""")

    ' Wrap in quotes
    web_Text = """" & web_Text & """"

    ' Escape !
    web_Text = VBA.Replace(web_Text, "!", """'!'""")

    ' Guard for ! at beginning or end (""'!'"..." or "..."'!'"" -> '!'"..." or "..."'!')
    If VBA.Left$(web_Text, 3) = """""'" Then
        web_Text = VBA.Right$(web_Text, VBA.Len(web_Text) - 2)
    End If
    If VBA.Right$(web_Text, 3) = "'""""" Then
        web_Text = VBA.Left$(web_Text, VBA.Len(web_Text) - 2)
    End If

    PrepareTextForShell = web_Text
End Function

''
' Prepare text for using with printf command
' - Wrap in "..."
' - Replace ! with '!' (reserved in bash)
' - Escape \, `, $, and "
' - Replace % with %% (used as an argument marker in printf)
'
' @internal
' @method PrepareTextForPrintf
' @param {String} Text
' @return {String}
''
Public Function PrepareTextForPrintf(ByVal web_Text As String) As String
    ' Escape special characters (except for !)
    web_Text = VBA.Replace(web_Text, "\", "\\")
    web_Text = VBA.Replace(web_Text, "`", "\`")
    web_Text = VBA.Replace(web_Text, "$", "\$")
    web_Text = VBA.Replace(web_Text, "%", "%%")
    web_Text = VBA.Replace(web_Text, """", "\""")

    ' Wrap in quotes
    web_Text = """" & web_Text & """"

    ' Escape !
    web_Text = VBA.Replace(web_Text, "!", """'!'""")

    ' Guard for ! at beginning or end (""'!'"..." or "..."'!'"" -> '!'"..." or "..."'!')
    If VBA.Left$(web_Text, 3) = """""'" Then
        web_Text = VBA.Right$(web_Text, VBA.Len(web_Text) - 2)
    End If
    If VBA.Right$(web_Text, 3) = "'""""" Then
        web_Text = VBA.Left$(web_Text, VBA.Len(web_Text) - 2)
    End If

    PrepareTextForPrintf = web_Text
End Function

' ============================================= '
' 8. Cryptography
' ============================================= '

''
' Determine the HMAC for the given text and secret using the SHA1 hash algorithm.
'
' Reference:
' - http://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac
'
' @example
' ```VB.net
' WebHelpers.HMACSHA1 "Howdy!", "Secret"
' ' -> c8fdf74a9d62aa41ac8136a1af471cec028fb157
' ```
'
' @method HMACSHA1
' @param {String} Text
' @param {String} Secret
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} HMAC-SHA1
''
Public Function HMACSHA1(Text As String, Secret As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -sha1 -hmac " & PrepareTextForShell(Secret)

    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If

    HMACSHA1 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_SecretBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)
    web_SecretBytes = VBA.StrConv(Secret, vbFromUnicode)

    ' Attempt to create system object. This will error if .NET Framework 3.5 is not available.
    On Error Resume Next
        Set web_Crypto = CreateObject("System.Security.Cryptography.HMACSHA1")
    On Error GoTo 0
    
    If web_Crypto Is Nothing Then
        ' If .NET Framework is unavailable, use WebCrypto class to perform hash.
        Set web_Crypto = New WebCrypto
        web_Crypto.InitHMAC web_SecretBytes
        web_Bytes = web_Crypto.HMACSHA1(web_TextBytes)
    Else
        web_Crypto.Key = web_SecretBytes
        web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)
    End If

    Select Case Format
    Case "Base64"
        HMACSHA1 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        HMACSHA1 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Determine the HMAC for the given text and secret using the SHA256 hash algorithm.
'
' @example
' ```VB.net
' WebHelpers.HMACSHA256 "Howdy!", "Secret"
' ' -> fb5d65...
' ```
'
' @method HMACSHA256
' @param {String} Text
' @param {String} Secret
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} HMAC-SHA256
''
Public Function HMACSHA256(Text As String, Secret As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -sha256 -hmac " & PrepareTextForShell(Secret)

    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If

    HMACSHA256 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_SecretBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)
    web_SecretBytes = VBA.StrConv(Secret, vbFromUnicode)

    ' Attempt to create system object. This will error if .NET Framework 3.5 is not available.
    On Error Resume Next
        Set web_Crypto = CreateObject("System.Security.Cryptography.HMACSHA256")
    On Error GoTo 0
    
    If web_Crypto Is Nothing Then
        ' If .NET Framework is unavailable, use WebCrypto class to perform hash.
        Set web_Crypto = New WebCrypto
        web_Crypto.InitHMAC web_SecretBytes
        web_Bytes = web_Crypto.HMACSHA256(web_TextBytes)
    Else
        web_Crypto.Key = web_SecretBytes
        web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)
    End If
    
    Select Case Format
    Case "Base64"
        HMACSHA256 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        HMACSHA256 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Determine the MD5 hash of the given text.
'
' Reference:
' - http://www.di-mgt.com.au/src/basMD5.bas.html
'
' @example
' ```VB.net
' WebHelpers.MD5 "Howdy!"
' ' -> 7105f32280940271293ee00ac97da5a7
' ```
'
' @method MD5
' @param {String} Text
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} MD5 Hash
''
Public Function MD5(Text As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -md5"

    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If

    MD5 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)

    ' Attempt to create system object. This will error if .NET Framework 3.5 is not available.
    On Error Resume Next
        Set web_Crypto = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    On Error GoTo 0
    
    If web_Crypto Is Nothing Then
        ' If .NET Framework is unavailable, use WebCrypto class to perform hash.
        Set web_Crypto = New WebCrypto
        web_Bytes = web_Crypto.MD5(web_TextBytes)
    Else
        web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)
    End If

    Select Case Format
    Case "Base64"
        MD5 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        MD5 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Determine the SHA256 hash of the given text.
'
' @method SHA256
' @param {String} Text
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} SHA256 Hash
''
Public Function SHA256(Text As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -sha256"
    
    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If
    
    SHA256 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, vbNullString)
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)

    ' Attempt to create system object. This will error if .NET Framework 3.5 is not available.
    On Error Resume Next
        'Set web_Crypto = CreateObject("System.Security.Cryptography.SHA256Managed")
    On Error GoTo 0
    
    If web_Crypto Is Nothing Then
        ' If .NET Framework is unavailable, use WebCrypto class to perform hash.
        Set web_Crypto = New WebCrypto
        web_Bytes = web_Crypto.SHA256(web_TextBytes)
    Else
        web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)
    End If

    Select Case Format
    Case "Base64"
        SHA256 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        SHA256 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Create random alphanumeric nonce (0-9a-zA-Z)
'
' @method CreateNonce
' @param {Integer} [NonceLength=32]
' @return {String} Randomly generated nonce
''
Public Function CreateNonce(Optional NonceLength As Integer = 32) As String
    Dim web_Str As String
    Dim web_Count As Integer
    Dim web_Result As String
    Dim web_Random As Integer

    web_Str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    web_Result = ""

    VBA.Randomize
    For web_Count = 1 To NonceLength
        web_Random = VBA.Int(((VBA.Len(web_Str) - 1) * VBA.Rnd) + 1)
        web_Result = web_Result & VBA.Mid$(web_Str, web_Random, 1)
    Next
    CreateNonce = web_Result
End Function

''
' Creates Challenge as required by PKCE authorisation flow.
'
' 1. SHA256 Hash is performed on Verifier.
' 2. SHA256 Hash (raw bytes, not hex-encoded string) is Base64 encoded.
' 3. Base64 encode modified to be Base64URL encoded.
'
' @method CreateChallenge
' @param {String} Verifier | A random string between 43 and 128 characters long that consists of the characters A-Z, a-z, 0-9, and the punctuation -._~ (hyphen, period, underscore, and tilde).
''
Public Function CreateChallenge(Verifier As String) As String
    CreateChallenge = VBA.Replace(VBA.Replace(VBA.Replace(web_AnsiBytesToBase64(web_SHA256Hash(Verifier)), "+", "-"), "/", "_"), "=", vbNullString)
End Function

''
' Convert string to ANSI bytes
'
' @internal
' @method StringToAnsiBytes
' @param {String} Text
' @return {Byte()}
''
Public Function StringToAnsiBytes(web_Text As String) As Byte()
    Dim web_Bytes() As Byte
    Dim web_AnsiBytes() As Byte
    Dim web_ByteIndex As Long
    Dim web_AnsiIndex As Long

    If VBA.Len(web_Text) > 0 Then
        ' Take first byte from unicode bytes
        ' VBA.Int is used for floor instead of round
        web_Bytes = web_Text
        ReDim web_AnsiBytes(VBA.Int(UBound(web_Bytes) / 2))

        web_AnsiIndex = LBound(web_Bytes)
        For web_ByteIndex = LBound(web_Bytes) To UBound(web_Bytes) Step 2
            web_AnsiBytes(web_AnsiIndex) = web_Bytes(web_ByteIndex)
            web_AnsiIndex = web_AnsiIndex + 1
        Next web_ByteIndex
    End If

    StringToAnsiBytes = web_AnsiBytes
End Function

#If Mac Then
#Else
Private Function web_AnsiBytesToBase64(web_Bytes() As Byte)
    ' Use XML to convert to Base64
    Dim web_XmlObj As Object
    Dim web_Node As Object

    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    Set web_Node = web_XmlObj.createElement("b64")

    web_Node.DataType = "bin.base64"
    web_Node.nodeTypedValue = web_Bytes
    web_AnsiBytesToBase64 = web_Node.Text

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
End Function

Private Function web_AnsiBytesToHex(web_Bytes() As Byte)
    Dim web_i As Long
    For web_i = LBound(web_Bytes) To UBound(web_Bytes)
        web_AnsiBytesToHex = web_AnsiBytesToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(web_Bytes(web_i)), 2))
    Next web_i
End Function
#End If

''
' SHA256 Hash of given Text. Returns hash as raw bytes (not hex-encoded string).
'
' @method web_SHA256Hash
' @param {String} Text
' @return {Byte}
''
Private Function web_SHA256Hash(Text As String) As Byte()
    Dim web_SHA256 As Object
    Dim web_TextToHash() As Byte
    
    web_TextToHash = StringToAnsiBytes(Text)
    
    ' Attempt to create system object. This will error if .NET Framework 3.5 is not available.
    On Error Resume Next
        Set web_SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    On Error GoTo 0
    
    If web_SHA256 Is Nothing Then
        ' If .NET Framework is unavailable, use WebCrypto class to perform hash.
        Set web_SHA256 = New WebCrypto
        web_SHA256Hash = web_SHA256.SHA256(web_TextToHash)
    Else
        web_SHA256Hash = web_SHA256.ComputeHash_2(web_TextToHash)
    End If
End Function

' ============================================= '
' 9. Converters
' ============================================= '

' Helper for url-encoded to create key=value pair
Private Function web_GetUrlEncodedKeyValue(Key As Variant, Value As Variant, Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.FormUrlEncoding) As String
    Select Case VBA.VarType(Value)
    Case VBA.vbBoolean
        ' Convert boolean to lowercase
        If Value Then
            Value = "true"
        Else
            Value = "false"
        End If
    Case VBA.vbDate
        ' Use region invariant date (ISO-8601)
        Value = ConvertToIso(CDate(Value))
    Case VBA.vbDecimal, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency
        ' Use region invariant number encoding ("." for decimal separator)
        Value = VBA.Replace(VBA.CStr(Value), ",", ".")
    End Select

    ' Url encode key and value (using + for spaces)
    web_GetUrlEncodedKeyValue = UrlEncode(Key, EncodingMode:=EncodingMode) & "=" & UrlEncode(Value, EncodingMode:=EncodingMode)
End Function

''
' AutoProxy 1.0.2
' (c) Damien Thirion
'
' Auto configure proxy server
'
' Based on code shared by Stephen Sulzer
' https://groups.google.com/d/msg/microsoft.public.winhttp/ZeWN2Xig82g/jgHIBDSfBwsJ
'
' Errors:
' 11020 - Unknown error while detecting proxy
' 11021 - WPAD detection failed
' 11022 - Unable to download proxy auto-config script
' 11023 - Error in proxy auto-config script
' 11024 - No proxy can be located for the specified URL
' 11025 - Specified URL is not valid
'
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Returns IE proxy settings
' including auto-detection and auto-config scripts results
'
' @param {String} Url
' @param[out] {String} ProxyServer
' @param[out] {String} ProxyBypass
''
Public Sub GetAutoProxy(ByVal Url As String, ByRef ProxyServer As String, ByRef ProxyBypass As String)
#If Mac Then
    ' (Windows only)
#ElseIf VBA7 Then
    Dim AutoProxy_ProxyStringPtr As LongPtr
    Dim AutoProxy_ptr As LongPtr
    Dim AutoProxy_hSession As LongPtr
#Else
    Dim AutoProxy_ProxyStringPtr As Long
    Dim AutoProxy_ptr As Long
    Dim AutoProxy_hSession As Long
#End If
#If Mac Then
#Else
    Dim AutoProxy_IEProxyConfig As AUTOPROXY_IE_PROXY_CONFIG
    Dim AutoProxy_AutoProxyOptions As AUTOPROXY_OPTIONS
    Dim AutoProxy_ProxyInfo As AUTOPROXY_INFO
    Dim AutoProxy_doAutoProxy As Boolean
    Dim AutoProxy_Error As Long
    Dim AutoProxy_ErrorMsg As String

    AutoProxy_AutoProxyOptions.AutoProxy_fAutoLogonIfChallenged = 1
    ProxyServer = ""
    ProxyBypass = ""

    ' WinHttpGetProxyForUrl returns unexpected errors if Url is empty
    If Url = "" Then Url = " "

    On Error GoTo AutoProxy_Cleanup

    ' Check IE's proxy configuration
    If (AutoProxy_GetIEProxy(AutoProxy_IEProxyConfig) > 0) Then
        ' If IE is configured to auto-detect, then we will too.
        If (AutoProxy_IEProxyConfig.AutoProxy_fAutoDetect <> 0) Then
            AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = AUTOPROXY_AUTO_DETECT
            AutoProxy_AutoProxyOptions.AutoProxy_dwAutoDetectFlags = _
                AUTOPROXY_DETECT_TYPE_DHCP + AUTOPROXY_DETECT_TYPE_DNS
            AutoProxy_doAutoProxy = True
        End If

        ' If IE is configured to use an auto-config script, then
        ' we will use it too
        If (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl <> 0) Then
            AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = _
                AutoProxy_AutoProxyOptions.AutoProxy_dwFlags + AUTOPROXY_CONFIG_URL
            AutoProxy_AutoProxyOptions.AutoProxy_lpszAutoConfigUrl = AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl
            AutoProxy_doAutoProxy = True
        End If
    Else
        ' If the IE proxy config is not available, then
        ' we will try auto-detection
        AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = AUTOPROXY_AUTO_DETECT
        AutoProxy_AutoProxyOptions.AutoProxy_dwAutoDetectFlags = _
            AUTOPROXY_DETECT_TYPE_DHCP + AUTOPROXY_DETECT_TYPE_DNS
        AutoProxy_doAutoProxy = True
    End If

    If AutoProxy_doAutoProxy Then
        On Error GoTo AutoProxy_TryIEFallback

        ' Need to create a temporary WinHttp session handle
        ' Note: Performance of this GetProxyInfoForUrl function can be
        '       improved by saving this AutoProxy_hSession handle across calls
        '       instead of creating a new handle each time
        AutoProxy_hSession = AutoProxy_HttpOpen(0, 1, 0, 0, 0)

        If (AutoProxy_GetProxyForUrl( _
            AutoProxy_hSession, StrPtr(Url), AutoProxy_AutoProxyOptions, AutoProxy_ProxyInfo) > 0) Then

            AutoProxy_ProxyStringPtr = AutoProxy_ProxyInfo.AutoProxy_lpszProxy
        Else
            AutoProxy_Error = Err.LastDllError
            Select Case AutoProxy_Error
            Case 12180
                AutoProxy_ErrorMsg = "WPAD detection failed"
                AutoProxy_Error = 10021
            Case 12167
                AutoProxy_ErrorMsg = "Unable to download proxy auto-config script"
                AutoProxy_Error = 10022
            Case 12166
                AutoProxy_ErrorMsg = "Error in proxy auto-config script"
                AutoProxy_Error = 10023
            Case 12178
                AutoProxy_ErrorMsg = "No proxy can be located for the specified URL"
                AutoProxy_Error = 10024
            Case 12005, 12006
                AutoProxy_ErrorMsg = "Specified URL is not valid"
                AutoProxy_Error = 10025
            Case Else
                AutoProxy_ErrorMsg = "Unknown error while detecting proxy"
                AutoProxy_Error = 10020
            End Select
        End If

        AutoProxy_HttpClose AutoProxy_hSession
        AutoProxy_hSession = 0
    End If

AutoProxy_TryIEFallback:
    On Error GoTo AutoProxy_Cleanup

    ' If we don't have a proxy server from WinHttpGetProxyForUrl,
    ' then pick one up from the IE proxy config (if given)
    If (AutoProxy_ProxyStringPtr = 0) Then
        AutoProxy_ProxyStringPtr = AutoProxy_IEProxyConfig.AutoProxy_lpszProxy
    End If

    ' If there's a proxy string, convert it to a Basic string
    If (AutoProxy_ProxyStringPtr <> 0) Then
        AutoProxy_ptr = AutoProxy_SysAllocString(AutoProxy_ProxyStringPtr)
        AutoProxy_CopyMemory VarPtr(ProxyServer), VarPtr(AutoProxy_ptr), 4
    End If

    ' Pick up any bypass string from the IEProxyConfig
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_ptr = AutoProxy_SysAllocString(AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass)
        AutoProxy_CopyMemory VarPtr(ProxyBypass), VarPtr(AutoProxy_ptr), 4
    End If

    ' Ensure WinHttp session is closed, an error might have occurred
    If (AutoProxy_hSession <> 0) Then
        AutoProxy_HttpClose AutoProxy_hSession
    End If

AutoProxy_Cleanup:
    On Error GoTo 0

    ' Free any strings received from WinHttp APIs
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl <> 0) Then
        AutoProxy_GlobalFree AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl
        AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl = 0
    End If
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxy <> 0) Then
        AutoProxy_GlobalFree AutoProxy_IEProxyConfig.AutoProxy_lpszProxy
        AutoProxy_IEProxyConfig.AutoProxy_lpszProxy = 0
    End If
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_GlobalFree AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass
        AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass = 0
    End If
    If (AutoProxy_ProxyInfo.AutoProxy_lpszProxy <> 0) Then
        AutoProxy_GlobalFree AutoProxy_ProxyInfo.AutoProxy_lpszProxy
        AutoProxy_ProxyInfo.AutoProxy_lpszProxy = 0
    End If
    If (AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_GlobalFree AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass
        AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass = 0
    End If

    ' Error handling
    If Err.Number <> 0 Then
        ' Unmanaged error
        Err.Raise Err.Number, "AutoProxy:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    ElseIf AutoProxy_Error <> 0 Then
        Err.Raise AutoProxy_Error, "AutoProxy", AutoProxy_ErrorMsg
    End If
#End If
End Sub
