Attribute VB_Name = "WebHelpers"
''
' WebHelpers v4.0.0-rc.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Common helpers VBA-Web
'
' Errors:
' 11000 - Error during parsing
' 11001 - Error during conversion
' 11002 - No matching converter has been registered
' 11003 - Error while getting url parts
' 11099 - (Temp) XML format is not currently supported on Mac
'
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
' VBA-JSON
' VBA-UTC
' AutoProxy
' --------------------------------------------- '

' === AutoProxy Headers
#If Mac Then
#ElseIf VBA7 Then
Private Declare PtrSafe Sub AutoProxy_CopyMemory Lib "kernel32" _
          Alias "RtlMoveMemory" (ByVal AutoProxy_lpDest As LongPtr, _
          ByVal AutoProxy_lpSource As LongPtr, ByVal AutoProxy_cbCopy As Long)
Private Declare PtrSafe Function AutoProxy_SysAllocString Lib "oleaut32" _
          Alias "SysAllocString" (ByVal AutoProxy_pwsz As LongPtr) As LongPtr
Private Declare PtrSafe Function AutoProxy_GlobalFree Lib "kernel32" _
          Alias "GlobalFree" (ByVal AutoProxy_p As LongPtr) As LongPtr
#Else
Private Declare Sub AutoProxy_CopyMemory Lib "kernel32" _
          Alias "RtlMoveMemory" (ByVal AutoProxy_lpDest As Long, _
          ByVal AutoProxy_lpSource As Long, ByVal AutoProxy_cbCopy As Long)
Private Declare Function AutoProxy_SysAllocString Lib "oleaut32" _
          Alias "GlobalFree" (ByVal AutoProxy_pwsz As Long) As Long
Private Declare Function AutoProxy_GlobalFree Lib "kernel32" _
          Alias "GlobalFree" (ByVal AutoProxy_p As Long) As Long
#End If
 
#If Mac Then
#ElseIf VBA7 Then
Private Declare PtrSafe Function AutoProxy_GetIEProxy Lib "WinHTTP.dll" _
     Alias "WinHttpGetIEProxyConfigForCurrentUser" _
    (ByRef AutoProxy_proxyConfig As AUTOPROXY_IE_PROXY_CONFIG) As Long
Private Declare PtrSafe Function AutoProxy_GetProxyForUrl Lib "WinHTTP.dll" _
    Alias "WinHttpGetProxyForUrl" _
    (ByVal AutoProxy_hSession As LongPtr, _
     ByVal AutoProxy_pszUrl As LongPtr, _
     ByRef AutoProxy_pAutoProxyOptions As AUTOPROXY_OPTIONS, _
     ByRef AutoProxy_pProxyInfo As AUTOPROXY_INFO) As Long
Private Declare PtrSafe Function AutoProxy_HttpOpen Lib "WinHTTP.dll" _
    Alias "WinHttpOpen" _
    (ByVal AutoProxy_pszUserAgent As LongPtr, _
     ByVal AutoProxy_dwAccessType As Long, _
     ByVal AutoProxy_pszProxyName As LongPtr, _
     ByVal AutoProxy_pszProxyBypass As LongPtr, _
     ByVal AutoProxy_dwFlags As Long) As LongPtr
Private Declare PtrSafe Function AutoProxy_HttpClose Lib "WinHTTP.dll" _
    Alias "WinHttpCloseHandle" _
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
Private Declare Function AutoProxy_GetIEProxy Lib "WinHTTP.dll" _
    Alias "WinHttpGetIEProxyConfigForCurrentUser" _
    (ByRef AutoProxy_proxyConfig As AUTOPROXY_IE_PROXY_CONFIG) As Long
Private Declare Function AutoProxy_GetProxyForUrl Lib "WinHTTP.dll" _
    Alias "WinHttpGetProxyForUrl" _
    (ByVal AutoProxy_hSession As Long, _
     ByVal AutoProxy_pszUrl As Long, _
     ByRef AutoProxy_pAutoProxyOptions As AUTOPROXY_OPTIONS, _
     ByRef AutoProxy_pProxyInfo As AUTOPROXY_INFO) As Long
Private Declare Function AutoProxy_HttpOpen Lib "WinHTTP.dll" _
    Alias "WinHttpOpen" _
    (ByVal AutoProxy_pszUserAgent As Long, _
     ByVal AutoProxy_dwAccessType As Long, _
     ByVal AutoProxy_pszProxyName As Long, _
     ByVal AutoProxy_pszProxyBypass As Long, _
     ByVal AutoProxy_dwFlags As Long) As Long
Private Declare Function AutoProxy_HttpClose Lib "WinHTTP.dll" _
        Alias "WinHttpCloseHandle" _
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
' ===

' === VBA-UTC Headers
#If Mac Then

Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" (ByVal utc_command As String, ByVal utc_mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" (ByVal utc_file As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" (ByVal utc_buffer As String, ByVal utc_size As Long, ByVal utc_number As Long, ByVal utc_file As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" (ByVal utc_file As Long) As Long

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#Else

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' ===

' === VBA-JSON Headers
#If Mac Then
#ElseIf Win64 Then
Private Declare PtrSafe Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)
#Else
Private Declare Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)
#End If
' ===

#If Mac Then
Private Declare Function web_popen Lib "libc.dylib" Alias "popen" (ByVal Command As String, ByVal mode As String) As Long
Private Declare Function web_pclose Lib "libc.dylib" Alias "pclose" (ByVal File As Long) As Long
Private Declare Function web_fread Lib "libc.dylib" Alias "fread" (ByVal outStr As String, ByVal size As Long, ByVal Items As Long, ByVal stream As Long) As Long
Private Declare Function web_feof Lib "libc.dylib" Alias "feof" (ByVal File As Long) As Long
#End If

Public Const WebUserAgent As String = "VBA-Web v4.0.0-rc.1 (https://github.com/VBA-tools/VBA-Web)"

' @internal
Public Type web_ShellResult
    web_Output As String
    web_ExitCode As Long
End Type

Private web_pDocumentHelper As Object
Private web_pElHelper As Object
Private web_pAsyncRequests As Dictionary
Private web_pConverters As Dictionary

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
    HttpGet = 0
    HttpPost = 1
    HttpPut = 2
    HttpDelete = 3
    HttpPatch = 4
End Enum
Public Enum WebFormat
    PlainText = 0
    Json = 1
    FormUrlEncoded = 2
    Xml = 3
    Custom = 9
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
            From = "VBA-Web"
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
Public Sub LogError(Message As String, Optional From As String = "", Optional ErrNumber As Long = 0)
    If From = "" Then
        From = "VBA-Web"
    End If
    If ErrNumber <> 0 Then
        If ErrNumber < 0 Then
            From = From & ": " & ErrNumber & " (" & VBA.LCase$(VBA.Hex$(ErrNumber)) & "), "
        Else
            From = From & ": " & ErrNumber & ", "
        End If
    Else
        From = From & ": "
    End If
    
    Debug.Print "ERROR - " & From & Message
End Sub

''
' Log request
'
' @param {WebRequest} Request
' --------------------------------------------- '
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
        
        If Request.Body <> "" Then
            Debug.Print vbNewLine & Request.Body
        End If
        
        Debug.Print
    End If
End Sub

''
' Log response
'
' @param {WebResponse} Response
' --------------------------------------------- '
Public Sub LogResponse(Client As WebClient, Request As WebRequest, Response As WebResponse)
    If EnableLogging Then
        Debug.Print "<-- Response - " & Format(Now, "Long Time")
        Debug.Print Response.StatusCode & " " & Response.StatusDescription
        
        Dim web_KeyValue As Dictionary
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
' Obfuscate message (for logging) by replacing with given character
'
' Example: ("Password", "#") -> ########
'
' @param {String} Secure
' @param {String} [Character = *]
' @return {String}
' --------------------------------------------- '
Public Function Obfuscate(Secure As String, Optional Character As String = "*") As String
    Obfuscate = VBA.String$(VBA.Len(Secure), Character)
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
' ParseJSON - Implemented in VBA-JSON embedded below

''
' Convert object to JSON string
'
' @param {Variant} Obj
' @return {String}
' --------------------------------------------- '
' ConvertToJSON - Implemented in VBA-JSON embedded below

''
' Parse url-encoded string to Dictionary
'
' @param {String} UrlEncoded
' @return {Dictionary} Parsed
' --------------------------------------------- '
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
' Convert Dictionary/Collection of key-value to url encoded string
'
' @param {Variant} Obj
' @return {String} UrlEncoded string (e.g. a=123&b=456&...)
' --------------------------------------------- '
Public Function ConvertToUrlEncoded(Obj As Variant) As String
    Dim web_Encoded As String

    If TypeOf Obj Is Collection Then
        Dim web_KeyValue As Dictionary
        For Each web_KeyValue In Obj
            If VBA.Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_KeyValue("Key"), web_KeyValue("Value"))
        Next web_KeyValue
    Else
        Dim web_Key As Variant
        For Each web_Key In Obj.Keys()
            If Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_Key, Obj(web_Key))
        Next web_Key
    End If
    
    ConvertToUrlEncoded = web_Encoded
End Function

''
' Parse XML string to XML
'
' @param {String} Encoded
' @return {Object} XML
' --------------------------------------------- '
Public Function ParseXML(Encoded As String) As Object
#If Mac Then
    LogError "XML format is not currently supported on Mac", "WebHelpers.ParseXML", 11050 + vbObjectError
    Err.Raise 11099, "WebHelpers.ParseXML", "XML format is not currently supported on Mac"
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
#If Mac Then
    LogError "XML format is not currently supported on Mac", "WebHelpers.ConvertToXML", 11050 + vbObjectError
    Err.Raise 11099 + vbObjectError, "WebHelpers.ConvertToXML", "XML format is not currently supported on Mac"
#Else
    On Error Resume Next
    ConvertToXML = Trim(Replace(Obj.Xml, vbCrLf, ""))
#End If
End Function

''
' Parse given string into object (Dictionary or Collection) for given format
'
' @param {String} Value
' @param {WebFormat} Format
' @param {String} [CustomFormat] Name of custom formatter to use
' @param {Variant} [Bytes] Raw response bytes to use with formatter (if specified)
' @return {Object|Dictionary|Collection}
' --------------------------------------------- '
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
        Set ParseByFormat = ParseXML(Value)
    Case WebFormat.Custom
        Dim web_Converter As Dictionary
        Dim web_Callback As String
        
        Set web_Converter = GetConverter(CustomFormat)
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
    End Select
    Exit Function
    
web_ErrorHandling:
    
    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during parsing" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
        
    LogError web_ErrorDescription, "WebHelpers.ParseByFormat", 11000 + vbObjectError
    Err.Raise 11000, "WebHelpers.ParseByFormat", web_ErrorDescription
End Function

''
' Convert object to given format
'
' @param {Variant|Dictionary|Collection} Obj
' @param {WebFormat} Format
' @param {String} [CustomFormat] Name of custom formatter to use
' @return {String}
' --------------------------------------------- '
Public Function ConvertToFormat(Obj As Variant, Format As WebFormat, Optional CustomFormat As String = "") As String
    On Error GoTo web_ErrorHandling

    Select Case Format
    Case WebFormat.Json
        ConvertToFormat = ConvertToJson(Obj)
    Case WebFormat.FormUrlEncoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case WebFormat.Xml
        ConvertToFormat = ConvertToXML(Obj)
    Case WebFormat.Custom
        Dim web_Converter As Dictionary
        Dim web_Callback As String
        
        Set web_Converter = GetConverter(CustomFormat)
        web_Callback = web_Converter("ConvertCallback")
        
        If web_Converter.Exists("web_Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter("web_Instance")
            ConvertToFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Obj)
        Else
            ConvertToFormat = Application.Run(web_Callback, Obj)
        End If
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
        
    LogError web_ErrorDescription, "WebHelpers.ConvertToFormat", 11001 + vbObjectError
    Err.Raise 11001, "WebHelpers.ConvertToFormat", web_ErrorDescription
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
        
        ' Set space value
        If SpaceAsPlus Then
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
                Case 36, 38, 43, 44, 47, 58, 59, 61, 63, 64
                    ' Reserved characters
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                Case 32, 34, 35, 37, 60, 62, 91 To 94, 96, 123 To 126
                    ' Unsafe characters
                    If EncodeUnsafe Then
                        If web_CharCode = 32 Then
                            web_Result(web_i) = web_Space
                        Else
                            web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                        End If
                    End If
                Case Else
                    web_Result(web_i) = web_Char
            End Select
        Next web_i
        UrlEncode = VBA.Join$(web_Result, "")
    End If
End Function

''
' Url decode the given encoded string
'
' @param {String} Encoded
' @return {String} Decoded string
' --------------------------------------------- '
Public Function UrlDecode(Encoded As String) As String
    Dim web_StringLen As Long
    web_StringLen = VBA.Len(Encoded)
    
    If web_StringLen > 0 Then
        Dim web_i As Long
        Dim web_Result As String
        Dim web_Temp As String
        
        For web_i = 1 To web_StringLen
            web_Temp = VBA.Mid$(Encoded, web_i, 1)
            
            If web_Temp = "+" Then
                web_Temp = " "
            ElseIf web_Temp = "%" And web_StringLen >= web_i + 2 Then
                web_Temp = VBA.Mid$(Encoded, web_i + 1, 2)
                web_Temp = VBA.Chr(VBA.CInt("&H" & web_Temp))
                
                web_i = web_i + 2
            End If
                
            web_Result = web_Result & web_Temp
        Next web_i
        
        UrlDecode = web_Result
    End If
End Function

''
' Url encode the given string
'
' @param {Variant} Text The raw string to encode
' @return {String} Encoded string
' --------------------------------------------- '
Public Function Base64Encode(Text As String) As String
    Base64Encode = VBA.Replace$(web_StringToBase64(Text), vbLf, "")
End Function

''
' Register custom converter to use with requests
'
' @param {String} Name
' @param {String} MediaType
' @param {String} ConvertCallback
' @param {String} ParseCallback
' @param {Object} [Instance]
' @param {String} [ParseType="String"] Use Content="String" or Body="Binary" in ParseCallback
' --------------------------------------------- '
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
Private Function GetConverter(CustomFormat As String) As Dictionary
    If web_pConverters.Exists(CustomFormat) Then
        Set GetConverter = web_pConverters(CustomFormat)
    Else
        LogError "No matching converter has been registered for custom format: " & CustomFormat, _
            "WebHelpers.GetConverter", 11002
        Err.Raise 11002, "WebHelpers.GetConverter", _
            "No matching converter has been registered for custom format: " & CustomFormat
    End If
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

    web_Results = Split(ExecuteInShell(web_Command).web_Output, " | ")
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

    Set UrlParts = web_Parts
    Exit Function
    
web_ErrorHandling:
    
    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while getting url parts" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
    
    LogError web_ErrorDescription, "WebHelpers.UrlParts", 11003
    Err.Raise 11003, "WebHelpers.UrlParts", web_ErrorDescription
End Function

' ============================================= '
' 4. Object/Dictionary/Collection/Array helpers
' ============================================= '

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
    Dim web_Key As Variant
    
    Set CloneDictionary = New Dictionary
    For Each web_Key In Dict.Keys
        CloneDictionary.Add VBA.CStr(web_Key), Dict(web_Key)
    Next web_Key
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
    Dim web_Item As Variant
    
    Set CloneCollection = New Collection
    For Each web_Item In Coll
        CloneCollection.Add web_Item
    Next web_Item
End Function

''
' Helper for creating key-value Dictionary for collection
'
' @param {String} Key
' @param {Variant} Value
' @return {Dictionary}
' --------------------------------------------- '
Public Function CreateKeyValue(Key As String, Value As Variant) As Dictionary
    Dim web_KeyValue As New Dictionary
    
    web_KeyValue("Key") = Key
    web_KeyValue("Value") = Value
    Set CreateKeyValue = web_KeyValue
End Function

''
' Helper for finding key-value in Collection of key-value
'
' @param {Collection} KeyValues
' @param {String} Key to find
' @return {Variant}
' --------------------------------------------- '
Public Function FindInKeyValues(KeyValues As Collection, Key As Variant) As Variant
    Dim web_KeyValue As Dictionary
    
    For Each web_KeyValue In KeyValues
        If web_KeyValue("Key") = Key Then
            FindInKeyValues = web_KeyValue("Value")
            Exit Function
        End If
    Next web_KeyValue
End Function

' ============================================= '
' 5. Request preparation / handling
' ============================================= '

''
' Get content-type for format
'
' @param {WebFormat} Format
' @return {String}
' --------------------------------------------- '
Public Function FormatToMediaType(Format As WebFormat, Optional CustomFormat As String) As String
    Select Case Format
    Case WebFormat.FormUrlEncoded
        FormatToMediaType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case WebFormat.Json
        FormatToMediaType = "application/json"
    Case WebFormat.Xml
        FormatToMediaType = "application/xml"
    Case WebFormat.Custom
        FormatToMediaType = GetConverter(CustomFormat)("MediaType")
    Case Else
        FormatToMediaType = "text/plain"
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
    End Select
End Function

''
' Add request to watched requests
'
' @internal
' @param {RestAsyncWrapper} AsyncWrapper
' --------------------------------------------- '
Public Sub AddAsyncRequest(AsyncWrapper As Object)
    If web_pAsyncRequests Is Nothing Then: Set web_pAsyncRequests = New Dictionary
    If Not AsyncWrapper.Request Is Nothing Then
        web_pAsyncRequests.Add AsyncWrapper.Request.Id, AsyncWrapper
    End If
End Sub

''
' Get watched request
'
' @internal
' @param {String} RequestId
' @return {RestAsyncWrapper}
' --------------------------------------------- '
Public Function GetAsyncRequest(RequestId As String) As Object
    If web_pAsyncRequests.Exists(RequestId) Then
        Set GetAsyncRequest = web_pAsyncRequests(RequestId)
    End If
End Function

''
' Remove request from watched requests
'
' @internal
' @param {String} RequestId
' --------------------------------------------- '
Public Sub RemoveAsyncRequest(RequestId As String)
    If Not web_pAsyncRequests Is Nothing Then
        If web_pAsyncRequests.Exists(RequestId) Then: web_pAsyncRequests.Remove RequestId
    End If
End Sub

' ============================================= '
' 6. Timing
' ============================================= '

''
' Start timeout timer for request
'
' @internal
' @param {WebRequest} Request
' @param {Long} TimeoutMS
' --------------------------------------------- '
Public Sub StartTimeoutTimer(AsyncWrapper As Object, TimeoutMs As Long)
    ' Round ms to seconds with minimum of 1 second if ms > 0
    Dim web_TimeoutS As Long
    web_TimeoutS = Round(TimeoutMs / 1000, 0)
    If TimeoutMs > 0 And web_TimeoutS = 0 Then
        web_TimeoutS = 1
    End If

    AddAsyncRequest AsyncWrapper
    Application.OnTime Now + TimeValue("00:00:" & web_TimeoutS), "'WebHelpers.TimeoutTimerExpired """ & AsyncWrapper.Request.Id & """'"
End Sub

''
' Stop timeout timer for request
'
' @internal
' @param {WebRequest} Request
' --------------------------------------------- '
Public Sub StopTimeoutTimer(AsyncWrapper As Object)
    If Not AsyncWrapper.Request Is Nothing Then
        RemoveAsyncRequest AsyncWrapper.Request.Id
    End If
End Sub

''
' Handle timeout timers expiring
'
' @internal
' @param {String} RequestId
' --------------------------------------------- '
Public Sub TimeoutTimerExpired(RequestId As String)
    Dim web_AsyncWrapper As Object
    Set web_AsyncWrapper = GetAsyncRequest(RequestId)
    
    If Not web_AsyncWrapper Is Nothing Then
        StopTimeoutTimer web_AsyncWrapper
        
        LogDebug "Async Timeout: " & web_AsyncWrapper.Request.FormattedResource, "WebHelpers.TimeoutTimerExpired"
        web_AsyncWrapper.TimedOut
    End If
End Sub

' ============================================= '
' 7. Mac
' ============================================= '
#If Mac Then

''
' Execute the given command
'
' @internal
' @param {String} Command
' @return {WebShellResult}
' --------------------------------------------- '
Public Function ExecuteInShell(Command As String) As web_ShellResult
    Dim web_File As Long
    Dim web_Chunk As String
    Dim web_Read As Long
    
    On Error GoTo web_Cleanup
    
    web_File = web_popen(Command, "r")
    
    If web_File = 0 Then
        ' TODO Investigate why this could happen and what should be done if it happens
        Exit Function
    End If
    
    Do While web_feof(web_File) = 0
        web_Chunk = VBA.Space$(50)
        web_Read = web_fread(web_Chunk, 1, Len(web_Chunk) - 1, web_File)
        If web_Read > 0 Then
            web_Chunk = VBA.Left$(web_Chunk, web_Read)
            ExecuteInShell.web_Output = ExecuteInShell.web_Output & web_Chunk
        End If
    Loop

web_Cleanup:

    ExecuteInShell.web_ExitCode = web_pclose(web_File)
End Function

''
' Prepare text for shell
' Wrap in "..." and replace ! with '!' (reserved in bash)
'
' @internal
' @param {String} Text
' @return {String}
' --------------------------------------------- '
Public Function PrepareTextForShell(ByVal Text As String) As String
    Text = VBA.Replace("""" & Text & """", "!", """'!'""")
    
    ' Guard for ! at beginning or end ("'!'"..." or "..."'!'" -> '!'"..." or "..."'!')
    If VBA.Left$(Text, 3) = """""'" Then
        Text = VBA.Right$(Text, VBA.Len(Text) - 2)
    End If
    If VBA.Right$(Text, 3) = "'""""" Then
        Text = VBA.Left$(Text, VBA.Len(Text) - 2)
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
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -sha1 -hmac " & PrepareTextForShell(Secret)
    HMACSHA1 = VBA.Replace(ExecuteInShell(web_Command).web_Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_Bytes() As Byte
    
    Set web_Crypto = CreateObject("System.Security.Cryptography.HMACSHA1")
    web_Crypto.Key = StringToANSIBytes(Secret)
    web_Bytes = web_Crypto.ComputeHash_2(StringToANSIBytes(Text))
    HMACSHA1 = web_ANSIBytesToHex(web_Bytes)
#End If

    If Format = "Base64" Then
        HMACSHA1 = web_HexToBase64(HMACSHA1)
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
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -sha256 -hmac " & PrepareTextForShell(Secret)
    HMACSHA256 = VBA.Replace(ExecuteInShell(web_Command).web_Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_Bytes() As Byte
    
    Set web_Crypto = CreateObject("System.Security.Cryptography.HMACSHA256")
    web_Crypto.Key = StringToANSIBytes(Secret)
    web_Bytes = web_Crypto.ComputeHash_2(StringToANSIBytes(Text))
    HMACSHA256 = web_ANSIBytesToHex(web_Bytes)
#End If

    If Format = "Base64" Then
        HMACSHA256 = web_HexToBase64(HMACSHA256)
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
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForShell(Text) & " | openssl dgst -md5"
    MD5 = VBA.Replace(ExecuteInShell(web_Command).web_Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_Bytes() As Byte
    
    Set web_Crypto = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    web_Bytes = web_Crypto.ComputeHash_2(StringToANSIBytes(Text))
    MD5 = web_ANSIBytesToHex(web_Bytes)
#End If

    If Format = "Base64" Then
        MD5 = web_HexToBase64(MD5)
    End If
End Function

''
' Create random alphanumeric nonce
'
' @param {Integer} [NonceLength=32]
' @return {String} Randomly generated nonce
' --------------------------------------------- '
Public Function CreateNonce(Optional NonceLength As Integer = 32) As String
    Dim web_Str As String
    Dim web_Count As Integer
    Dim web_Result As String
    Dim web_Random As Integer
    
    web_Str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    web_Result = ""
    
    For web_Count = 1 To NonceLength
        web_Random = VBA.Int(((VBA.Len(web_Str) - 1) * VBA.Rnd) + 1)
        web_Result = web_Result + VBA.Mid$(web_Str, web_Random, 1)
    Next
    CreateNonce = web_Result
End Function

Public Function StringToANSIBytes(Text As String) As Byte()
    Dim web_Bytes() As Byte
    Dim web_ANSIBytes() As Byte
    Dim web_ByteIndex As Long
    Dim web_ANSIIndex As Long
    
    If VBA.Len(Text) > 0 Then
        ' Take first byte from unicode bytes
        ' VBA.Int is used for floor instead of round
        web_Bytes = Text
        ReDim web_ANSIBytes(VBA.Int(UBound(web_Bytes) / 2))
        
        web_ANSIIndex = LBound(web_Bytes)
        For web_ByteIndex = LBound(web_Bytes) To UBound(web_Bytes) Step 2
            web_ANSIBytes(web_ANSIIndex) = web_Bytes(web_ByteIndex)
            web_ANSIIndex = web_ANSIIndex + 1
        Next web_ByteIndex
    End If
    
    StringToANSIBytes = web_ANSIBytes
End Function

Private Function web_ANSIBytesToString(web_Bytes() As Byte) As String
    Dim web_i As Long
    For web_i = LBound(web_Bytes) To UBound(web_Bytes)
        web_ANSIBytesToString = web_ANSIBytesToString & VBA.Chr$(web_Bytes(web_i))
    Next web_i
End Function

Private Function web_HexToANSIBytes(web_Hex As String) As Byte()
    Dim web_Bytes() As Byte
    Dim web_HexIndex As Integer
    Dim web_ByteIndex As Integer

    ' Remove linefeeds
    web_Hex = VBA.Replace(web_Hex, vbLf, "")

    ReDim web_Bytes(VBA.Len(web_Hex) / 2 - 1)
    web_ByteIndex = 0
    For web_HexIndex = 1 To Len(web_Hex) Step 2
        web_Bytes(web_ByteIndex) = VBA.CLng("&H" & VBA.Mid$(web_Hex, web_HexIndex, 2))
        web_ByteIndex = web_ByteIndex + 1
    Next web_HexIndex
    
    web_HexToANSIBytes = web_Bytes
End Function

Private Function web_ANSIBytesToHex(web_Bytes() As Byte)
    Dim web_i As Long
    For web_i = LBound(web_Bytes) To UBound(web_Bytes)
        web_ANSIBytesToHex = web_ANSIBytesToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(web_Bytes(web_i)), 2))
    Next web_i
End Function

Private Function web_StringToHex(web_Text As String) As String
    ' Convert single-byte character to hex
    ' (May need better international handling in the future)
    Dim web_Bytes() As Byte
    Dim web_i As Integer
    
    web_Bytes = StringToANSIBytes(web_Text)
    For web_i = LBound(web_Bytes) To UBound(web_Bytes)
        web_StringToHex = web_StringToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(web_Bytes(web_i)), 2))
    Next web_i
End Function

Private Function web_StringToBase64(ByVal web_Text As String) As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForShell(web_Text) & " | openssl base64"
    web_StringToBase64 = ExecuteInShell(web_Command).web_Output
#Else
    ' Use XML to convert to Base64
    ' but XML requires bytes, so convert to bytes first
    Dim web_XmlObj As Object
    Dim web_Node As Object
    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    
    Set web_Node = web_XmlObj.createElement("b64")
    web_Node.DataType = "bin.base64"
    web_Node.nodeTypedValue = StringToANSIBytes(web_Text)
    web_StringToBase64 = web_Node.Text

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
#End If
End Function

Private Function web_HexToBase64(ByVal web_Hex As String) As String
    web_HexToBase64 = web_StringToBase64(web_ANSIBytesToString(web_HexToANSIBytes(web_Hex)))
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
Private Function web_GetUrlEncodedKeyValue(Key As Variant, Value As Variant) As String
    ' Convert boolean to lowercase
    If VBA.VarType(Value) = VBA.vbBoolean Then
        If Value Then
            Value = "true"
        Else
            Value = "false"
        End If
    End If
    
    ' Url encode key and value (using + for spaces)
    web_GetUrlEncodedKeyValue = UrlEncode(Key, SpaceAsPlus:=True) & "=" & UrlEncode(Value, SpaceAsPlus:=True)
End Function

''
' VBA-JSON v1.0.0-rc.1
' (c) Tim Hall - https://github.com/timhall/VBA-JSONConverter
'
' JSON Converter for VBA
'
' Errors (513-65535 available):
' 10001 - JSON parse error
' 10002 - ISO 8601 date conversion error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
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

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' -------------------------------------- '
Public Function ParseJson(ByVal json_String As String, Optional json_ConvertLargeNumbersToString As Boolean = True) As Object
    Dim json_Index As Long
    json_Index = 1
    
    ' Remove vbCr, vbLf, and vbTab from json_String
    json_String = VBA.Replace(VBA.Replace(VBA.Replace(json_String, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(json_String, json_Index, json_ConvertLargeNumbersToString)
    Case "["
        Set ParseJson = json_ParseArray(json_String, json_Index, json_ConvertLargeNumbersToString)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @param {Variant} json_DictionaryCollectionOrArray (Dictionary, Collection, or Array)
' @return {String}
' -------------------------------------- '
Public Function ConvertToJson(ByVal json_DictionaryCollectionOrArray As Variant, Optional json_ConvertLargeNumbersFromString As Boolean = True) As String
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    
    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True

    Select Case VBA.VarType(json_DictionaryCollectionOrArray)
    Case VBA.vbNull, VBA.vbEmpty
        ConvertToJson = "null"
    Case VBA.vbDate
        ' Date
        json_DateStr = ConvertToIso(VBA.CDate(json_DictionaryCollectionOrArray))
        
        ConvertToJson = """" & json_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If json_ConvertLargeNumbersFromString And json_StringIsLargeNumber(json_DictionaryCollectionOrArray) Then
            ConvertToJson = json_DictionaryCollectionOrArray
        Else
            ConvertToJson = """" & json_Encode(json_DictionaryCollectionOrArray) & """"
        End If
    Case VBA.vbBoolean
        If json_DictionaryCollectionOrArray Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        ' Array
        json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
        
        On Error Resume Next
        
        json_LBound = LBound(json_DictionaryCollectionOrArray, 1)
        json_UBound = UBound(json_DictionaryCollectionOrArray, 1)
        json_LBound2D = LBound(json_DictionaryCollectionOrArray, 2)
        json_UBound2D = UBound(json_DictionaryCollectionOrArray, 2)
        
        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_Index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                End If
            
                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
                
                    For json_Index2D = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                        End If
                        
                        json_BufferAppend json_buffer, _
                            ConvertToJson(json_DictionaryCollectionOrArray(json_Index, json_Index2D), _
                                json_ConvertLargeNumbersFromString), _
                            json_BufferPosition, json_BufferLength
                    Next json_Index2D
                    
                    json_BufferAppend json_buffer, "]", json_BufferPosition, json_BufferLength
                    json_IsFirstItem2D = True
                Else
                    json_BufferAppend json_buffer, _
                        ConvertToJson(json_DictionaryCollectionOrArray(json_Index), _
                            json_ConvertLargeNumbersFromString), _
                        json_BufferPosition, json_BufferLength
                End If
            Next json_Index
        End If
        
        On Error GoTo 0
        
        json_BufferAppend json_buffer, "]", json_BufferPosition, json_BufferLength
        
        ConvertToJson = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
    
    ' Dictionary or Collection
    Case VBA.vbObject
        ' Dictionary
        If VBA.TypeName(json_DictionaryCollectionOrArray) = "Dictionary" Then
            json_BufferAppend json_buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In json_DictionaryCollectionOrArray.Keys
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                End If
            
                json_BufferAppend json_buffer, _
                    """" & json_Key & """:" & ConvertToJson(json_DictionaryCollectionOrArray(json_Key), json_ConvertLargeNumbersFromString), _
                    json_BufferPosition, json_BufferLength
            Next json_Key
            json_BufferAppend json_buffer, "}", json_BufferPosition, json_BufferLength
        
        ' Collection
        ElseIf VBA.TypeName(json_DictionaryCollectionOrArray) = "Collection" Then
            json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In json_DictionaryCollectionOrArray
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                End If
            
                json_BufferAppend json_buffer, _
                    ConvertToJson(json_Value, json_ConvertLargeNumbersFromString), _
                    json_BufferPosition, json_BufferLength
            Next json_Value
            json_BufferAppend json_buffer, "]", json_BufferPosition, json_BufferLength
        End If
        
        ConvertToJson = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
    Case Else
        ' Number
        On Error Resume Next
        ConvertToJson = json_DictionaryCollectionOrArray
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Dictionary
    Dim json_Key As String
    Dim json_NextChar As String
    
    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1
        
        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If
            
            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index, json_ConvertLargeNumbersToString)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index, json_ConvertLargeNumbersToString)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Collection
    Set json_ParseArray = New Collection
    
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1
        
        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If
            
            json_ParseArray.Add json_ParseValue(json_String, json_Index, json_ConvertLargeNumbersToString)
        Loop
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index, json_ConvertLargeNumbersToString)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    
    json_SkipSpaces json_String, json_Index
    
    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1
    
    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)
        
        Select Case json_Char
        Case "\"
            ' Escaped string, \\, or \/
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            
            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Variant
    Dim json_Char As String
    Dim json_Value As String
    
    json_SkipSpaces json_String, json_Index
    
    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)
        
        If VBA.InStr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15 characters containing only numbers and decimal points -> Number
            If json_ConvertLargeNumbersToString And Len(json_Value) >= 16 Then
                json_ParseNumber = json_Value
            Else
                ' Guard for regional settings that use "," for decimal
                ' CStr(0.1) -> "0.1" or "0,1" based on regional settings -> Replace "." with "." or ","
                json_Value = VBA.Replace(json_Value, ".", VBA.Mid$(VBA.CStr(0.1), 2, 1))
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    ' Parse key with single or double quotes
    json_ParseKey = json_ParseString(json_String, json_Index)
    
    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    
    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)
        
        Select Case json_AscCode
        ' " -> 34 -> \"
        Case 34
            json_Char = "\"""
        ' \ -> 92 -> \\
        Case 92
            json_Char = "\\"
        ' / -> 47 -> \/
        Case 47
            json_Char = "\/"
        ' backspace -> 8 -> \b
        Case 8
            json_Char = "\b"
        ' form feed -> 12 -> \f
        Case 12
            json_Char = "\f"
        ' line feed -> 10 -> \n
        Case 10
            json_Char = "\n"
        ' carriage return -> 13 -> \r
        Case 13
            json_Char = "\r"
        ' tab -> 9 -> \t
        Case 9
            json_Char = "\t"
        ' Non-ascii characters -> convert to 4-digit hex
        Case 0 To 31, 127 To 65535
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select
            
        json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index
    
    json_Encode = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)
    
    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String
        Dim json_Index As Long
        
        json_StringIsLargeNumber = True
        
        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['
    
    Dim json_StartIndex As Long
    Dim json_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
#If Mac Then
    json_buffer = json_buffer & json_Append
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

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long
    
    json_AppendLength = VBA.LenB(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition
    
    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim json_TemporaryLength As Long
        
        json_TemporaryLength = json_BufferLength
        Do While json_TemporaryLength < json_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If json_TemporaryLength = 0 Then
                json_TemporaryLength = json_TemporaryLength + 510
            Else
                json_TemporaryLength = json_TemporaryLength + 16384
            End If
        Loop
        
        json_buffer = json_buffer & VBA.Space$((json_TemporaryLength - json_BufferLength) \ 2)
        json_BufferLength = json_TemporaryLength
    End If
    
    ' Copy memory from append to buffer at buffer position
    json_CopyMemory ByVal json_UnsignedAdd(StrPtr(json_buffer), _
                    json_BufferPosition), _
                    ByVal StrPtr(json_Append), _
                    json_AppendLength
    
    json_BufferPosition = json_BufferPosition + json_AppendLength
#End If
End Sub

Private Function json_BufferToString(ByRef json_buffer As String, ByVal json_BufferPosition As Long, ByVal json_BufferLength As Long) As String
#If Mac Then
    json_BufferToString = json_buffer
#Else
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_buffer, json_BufferPosition \ 2)
    End If
#End If
End Function

#If Win64 Then
Private Function json_UnsignedAdd(json_Start As LongPtr, json_Increment As Long) As LongPtr
#Else
Private Function json_UnsignedAdd(json_Start As Long, json_Increment As Long) As Long
#End If

    If json_Start And &H80000000 Then
        json_UnsignedAdd = json_Start + json_Increment
    ElseIf (json_Start Or &H80000000) < -json_Increment Then
        json_UnsignedAdd = json_Start + json_Increment
    Else
        json_UnsignedAdd = (json_Start + &H80000000) + (json_Increment + &H80000000)
    End If
End Function

''
' VBA-UTC v1.0.0-rc.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @param {Date} utc_UtcDate
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo web_ErrorHandling
    
#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME
    
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate
    
    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

web_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' -------------------------------------- '
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo web_ErrorHandling
    
#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME
    
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate
    
    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If
    
    Exit Function
    
web_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @param {Date} utc_IsoString
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo web_ErrorHandling
    
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date
    
    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
    
    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If
            
            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")
                
                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), VBA.CInt(utc_OffsetParts(2)))
                End Select
                
                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If
        
        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), VBA.CInt(utc_TimeParts(2)))
        End Select
        
        If utc_HasOffset Then
            ParseIso = ParseIso + utc_Offset
        Else
            ParseIso = ParseUtc(ParseIso)
        End If
    End If
    
    Exit Function
    
web_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' -------------------------------------- '
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo web_ErrorHandling
    
    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
    
    Exit Function
    
web_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then
Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    
    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If
    
    utc_Result = utc_ExecuteInShell(utc_ShellCommand)
    
    If utc_Result.utc_Output = "" Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = Split(utc_Result.utc_Output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")
        
        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function
Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
    Dim utc_file As Long
    Dim utc_Chunk As String
    Dim utc_Read As Long
    
    On Error GoTo web_ErrorHandling
    utc_file = utc_popen(utc_ShellCommand, "r")
    
    If utc_file = 0 Then: Exit Function
    
    Do While utc_feof(utc_file) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_file)
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, utc_Read)
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

web_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = utc_pclose(File)
End Function
#Else
Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function
#End If

''
' AutoProxy 1.0
'
' (c) Damien Thirion
' Based on code shared by Stephen Sulzer
' https://groups.google.com/d/msg/microsoft.public.winhttp/ZeWN2Xig82g/jgHIBDSfBwsJ
'
' Auto configure proxy server
'
' Returns IE proxy settings
' including auto-detection and auto-config scripts results
'
'
' @param {String} Url
' @param[out] {String} ProxyServer
' @param[out] {String} ProxyBypass
' --------------------------------------------- '
Public Sub GetAutoProxy(ByVal Url As String, ByRef ProxyServer As String, ByRef ProxyBypass As String)
#If Mac Then
    ' Windows only !
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
                        AUTOPROXY_DETECT_TYPE_DHCP + _
                        AUTOPROXY_DETECT_TYPE_DNS
            AutoProxy_doAutoProxy = True
        End If
     
        ' If IE is configured to use an auto-config script, then
        ' we will use it too
        If (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl <> 0) Then
            AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = AutoProxy_AutoProxyOptions.AutoProxy_dwFlags + _
                        AUTOPROXY_CONFIG_URL
            AutoProxy_AutoProxyOptions.AutoProxy_lpszAutoConfigUrl = AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl
            AutoProxy_doAutoProxy = True
        End If
    Else
        ' if the IE proxy config is not available, then
        ' we will try auto-detection
        AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = AUTOPROXY_AUTO_DETECT
        AutoProxy_AutoProxyOptions.AutoProxy_dwAutoDetectFlags = _
                        AUTOPROXY_DETECT_TYPE_DHCP + _
                        AUTOPROXY_DETECT_TYPE_DNS
        AutoProxy_doAutoProxy = True
    End If
    
    If AutoProxy_doAutoProxy Then
        On Error GoTo AutoProxy_TryIEFallback
        
        ' Need to create a temporary WinHttp session handle
        '  Note: performance of this GetProxyInfoForUrl function can be
        '   improved by saving this AutoProxy_hSession handle across calls
        '   instead of creating a new handle each time
        AutoProxy_hSession = AutoProxy_HttpOpen(0, 1, 0, 0, 0)
     
        If (AutoProxy_GetProxyForUrl(AutoProxy_hSession, StrPtr(Url), AutoProxy_AutoProxyOptions, _
                AutoProxy_ProxyInfo) > 0) Then
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
     
        AutoProxy_HttpClose (AutoProxy_hSession)
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
        AutoProxy_HttpClose (AutoProxy_hSession)
    End If
    
AutoProxy_Cleanup:
    ' Free any strings received from WinHttp APIs
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl <> 0) Then
        AutoProxy_GlobalFree (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl)
    End If
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxy <> 0) Then
        AutoProxy_GlobalFree (AutoProxy_IEProxyConfig.AutoProxy_lpszProxy)
    End If
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_GlobalFree (AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass)
    End If
    If (AutoProxy_ProxyInfo.AutoProxy_lpszProxy <> 0) Then
        AutoProxy_GlobalFree (AutoProxy_ProxyInfo.AutoProxy_lpszProxy)
    End If
    If (AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_GlobalFree (AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass)
    End If
    
    ' Error handling
    If Err <> 0 Then
        ' Unmanaged error
        Err.Raise Err.Number, "AutoProxy:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    ElseIf AutoProxy_Error <> 0 Then
        On Error GoTo 0
        Err.Raise AutoProxy_Error, "AutoProxy", AutoProxy_ErrorMsg
    End If
#End If
End Sub
