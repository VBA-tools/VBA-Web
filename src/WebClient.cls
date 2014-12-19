VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebClient v4.0.0-rc.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Interact with REST web services from Excel
'
' Errors:
' 11010 / 80042b02 / -2147210494 - cURL error in Execute
' 11011 / 80042b03 / -2147210493 - Error in Execute
' 11012 / 80042b04 / -2147210492 - Error preparing http request
' 11013 / 80042b05 / -2147210491 - Error preparing cURL request
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const web_DefaultTimeoutMs As Long = 5000

Private pBaseUrl As String
Private pProxyServer As String
Private pAutoProxyLoaded As Boolean

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Authenticator As IWebAuthenticator
Public TimeoutMs As Long
Public Username As String
Public Password As String
Public ProxyUsername As String
Public ProxyPassword As String
Public ProxyBypassList As String
Public EnableAutoProxy As Boolean

Public Property Get ProxyServer() As String
    If Me.EnableAutoProxy And Not pAutoProxyLoaded Then
        LogDebug "Retrieving AutoProxyServer"
        On Error Resume Next
        pProxyServer = WebHelpers.GetAutoProxyServer(Me.BaseUrl)
        If Err <> 0 Then
            LogDebug "AutoProxyServer: " & Err.Description
        End If
        On Error GoTo 0
        pAutoProxyLoaded = True
    End If

    ProxyServer = pProxyServer
End Property

Public Property Let ProxyServer(Value As String)
    ' User defined manually a Proxy -> won't try AutoProxy
    pProxyServer = Value
    Me.EnableAutoProxy = False
End Property

Public Property Get BaseUrl() As String
    BaseUrl = pBaseUrl
End Property

Public Property Let BaseUrl(Value As String)
    pBaseUrl = Value
    ' A new BaseUrl was defined. Flag the AutoProxy to be reevaluated
    pAutoProxyLoaded = False
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Execute the specified request
'
' @param {WebRequest} request The request to execute
' @return {WebResponse} Wrapper of server response for request
' --------------------------------------------- '
Public Function Execute(Request As WebRequest) As WebResponse
    Dim web_Http As Object
    Dim web_Response As New WebResponse
    
    On Error GoTo web_ErrorHandling
    
#If Mac Then
    Dim web_Curl As String
    Dim web_Result As web_ShellResult
    
    web_Curl = Me.PrepareCurlRequest(Request)
    web_Result = WebHelpers.ExecuteInShell(web_Curl)
    
    If web_Result.web_ExitCode > 0 Then
        Dim web_ErrorNumber As Long
        Dim web_ErrorMessage As String
        web_ErrorNumber = web_Result.web_ExitCode / 256
        
        Select Case web_ErrorNumber
        Case 1
            ' 1 = CURLE_UNSUPPORTED_PROTOCOL
            Err.Raise 208614 + vbObjectError, "The URL does not use a recognized protocol (1: CURLE_UNSUPPORTED_PROTOCOL)" & vbNewLine & _
                "URL: " & Me.GetFullUrl(Request) & vbNewLine & _
                "Protocol: " & WebHelpers.UrlParts(Me.GetFullUrl(Request))("Protocol")
        Case 3
            ' 3 = CURLE_URL_MALFORMAT
            Err.Raise 208613 + vbObjectError, "The URL is invalid (3: CURLE_URL_MALFORMAT)" & _
                "URL: " & Me.GetFullUrl(Request)
        Case 5, 6, 7
            ' 5 = CURLE_COULDNT_RESOLVE_PROXY
            ' 6 = CURLE_COULDNT_RESOLVE_HOST
            ' 7 = CURLE_COULDNT_CONNECT
            Err.Raise 208637 + vbObjectError, "WebClient.Execute", "A connection with the server could not be established"
        Case 12, 28
            ' 12 = CURLE_FTP_ACCEPT_TIMEOUT
            ' 28 = CURLE_OPERATION_TIMEDOUT
            Err.Raise 208610 + vbObjectError, "WebClient.Execute", "The operation timed out"
        Case Else
            Err.Raise 11010 + vbObjectError, "WebClient.Execute", "An unknown cURL error occured, #" & web_ErrorNumber & vbNewLine & _
                "Find details at http://curl.haxx.se/libcurl/c/libcurl-errors.html"
        End Select
    End If
    
    web_Response.CreateFromCurl Me, Request, web_Result.web_Output

#Else
    Set web_Http = Me.PrepareHttpRequest(Request)
    
    web_Http.Send Request.Body
    Do While Not web_Http.WaitForResponse(25)
        VBA.DoEvents
    Loop
    
    web_Response.CreateFromHttp Me, Request, web_Http
    
#End If

    WebHelpers.LogResponse Me, Request, web_Response
    
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.AfterExecute Me, Request, web_Response
    End If
    
    Set web_Http = Nothing
    Set Execute = web_Response
    Exit Function
    
web_ErrorHandling:

    Set web_Http = Nothing
    Dim web_ErrorDescription As String
    
    Select Case Err.Number - vbObjectError
    Case 208610, 208637
        ' Return 408
        web_Response.StatusCode = WebStatusCode.RequestTimeout
        web_Response.StatusDescription = "Request Timeout"
        
        WebHelpers.LogResponse Me, Request, web_Response
        Set Execute = web_Response
        Err.Clear
    Case 11010
        ' cURL error
        WebHelpers.LogError Err.Description, "WebClient.Execute", Err.Number
        Err.Raise Err.Number, "WebClient.Execute", Description:=Err.Description
    Case Else
        ' Error
        web_ErrorDescription = "An error occurred during execute" & vbNewLine & _
            Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
        
        WebHelpers.LogError web_ErrorDescription, "WebClient.Execute", 11011 + vbObjectError
        Err.Raise 11011 + vbObjectError, "WebClient.Execute", web_ErrorDescription
    End Select
End Function

''
' GET JSON from Url using options
'
' @param {String} Url (relative to BaseUrl, if set)
' @param {Dictionary} [Options]
' - Headers
' - Cookies
' - QuerystringParams
' - UrlSegments
' @return {WebResponse} Response
' --------------------------------------------- '
Public Function GetJson(Url As String, Optional Options As Dictionary) As WebResponse
    Dim web_Request As New WebRequest
    web_Request.CreateFromOptions Options
    web_Request.Resource = Url
    web_Request.Format = WebFormat.Json
    web_Request.Method = WebMethod.HttpGet
    
    Set GetJson = Me.Execute(web_Request)
End Function

''
' POST JSON to Url using body and options
'
' @param {String} Url (relative to BaseUrl, if set)
' @param {Dictionary} Body
' @param {Dictionary} [Options]
' - Headers
' - Cookies
' - QuerystringParams
' - UrlSegments
' @return {WebResponse} Response
' --------------------------------------------- '
Public Function PostJson(Url As String, Body As Variant, Optional Options As Dictionary) As WebResponse
    Dim web_Request As New WebRequest
    web_Request.CreateFromOptions Options
    web_Request.Resource = Url
    web_Request.Format = WebFormat.Json
    web_Request.Method = WebMethod.HttpPost
    If IsObject(Body) Then
        Set web_Request.Body = Body
    Else
        web_Request.Body = Body
    End If
    
    Set PostJson = Me.Execute(web_Request)
End Function

''
' Set proxy for all requests
'
' @param {String} ProxyServer
' @param {String} [Username=""]
' @param {String} [Password=""]
' @param {Variant} [BypassList]
' --------------------------------------------- '
Public Sub SetProxy(ProxyServer As String, _
    Optional Username As String = "", Optional Password As String = "", Optional BypassList As String = "")
    
    Me.ProxyServer = ProxyServer
    Me.ProxyUsername = Username
    Me.ProxyPassword = Password
    Me.ProxyBypassList = BypassList
End Sub

''
' Prepare Http request for given WebRequest
'
' @internal
' @param {WebRequest} Request
' @return {WinHttpRequest}
' --------------------------------------------- '
Public Function PrepareHttpRequest(Request As WebRequest, Optional Async As Boolean = True) As Object
    Dim web_Http As Object
    
    On Error GoTo web_ErrorHandling
    
    Set web_Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Set timeouts
    web_Http.SetTimeouts Me.TimeoutMs, Me.TimeoutMs, Me.TimeoutMs, Me.TimeoutMs
    
    ' Setup authenticator
    web_BeforeExecute Request
    
    ' Open http request
    web_Http.Open WebHelpers.MethodToName(Request.Method), Me.GetFullUrl(Request), Async
    
    ' Setup proxy
    If Me.ProxyServer <> "" Then
        LogDebug "SetProxy: " & Me.ProxyServer, "WebClient.PrepareHttpRequest"
        web_Http.SetProxy 2, Me.ProxyServer, Me.ProxyBypassList
        
        If Me.ProxyUsername <> "" Then
            LogDebug "SetProxyCredentials: " & Me.ProxyUsername & ", " & WebHelpers.Obfuscate(Me.ProxyPassword), "WebClient.PrepareHttpRequest"
            web_Http.SetCredentials Me.ProxyUsername, Me.ProxyPassword, 1 ' 1 = Credentials for proxy
        End If
    Else
        ' Attempt to get proxy setup with Proxycfg.exe, otherwise direct
        ' See http://msdn.microsoft.com/en-us/library/windows/desktop/aa384059(v=vs.85).aspx for details
        web_Http.SetProxy 1 ' 1 = HTTPREQUEST_PROXYSETTING_PRECONFIG
    End If
    
    ' Setup Basic authentication
    If Me.Username <> "" Then
        web_Http.SetCredentials Me.Username, Me.Password, 0 ' 0 = Credentials for server
        Request.AddHeader "Authorization", "Basic " & WebHelpers.Base64Encode(Me.Username & ":" & Me.Password)
    End If
    
    ' Set headers on http request (after open)
    Dim web_KeyValue As Dictionary
    For Each web_KeyValue In Request.Headers
        web_Http.setRequestHeader web_KeyValue("Key"), web_KeyValue("Value")
    Next web_KeyValue
    
    For Each web_KeyValue In Request.Cookies
        web_Http.setRequestHeader "Cookie", web_KeyValue("Key") & "=" & web_KeyValue("Value")
    Next web_KeyValue
    
    ' Log request and return
    WebHelpers.LogRequest Me, Request
    Set PrepareHttpRequest = web_Http
    Exit Function
    
web_ErrorHandling:

    Set web_Http = Nothing
    Err.Raise 11012 + vbObjectError, "WebClient.PrepareHttpRequest", _
        "An error occurred while preparing http request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

''
' Prepare cURL request for given WebRequest
'
' @internal
' @param {WebRequest} Request
' @return {String}
' --------------------------------------------- '
Public Function PrepareCurlRequest(Request As WebRequest) As String
    Dim web_Curl As String
    
    On Error GoTo web_ErrorHandling
    
    web_Curl = "curl -i"
    
    ' Set timeouts
    ' (max time = resolve + sent + receive)
    web_Curl = web_Curl & " --connect-timeout " & Me.TimeoutMs / 1000
    web_Curl = web_Curl & " --max-time " & 3 * Me.TimeoutMs / 1000
    
    ' Setup proxy
    If Me.ProxyServer <> "" Then
        web_Curl = web_Curl & " --proxy " & Me.ProxyServer
        
        If Me.ProxyBypassList <> "" Then
            web_Curl = web_Curl & " --noproxy " & Me.ProxyBypassList
        End If
        If Me.ProxyUsername <> "" Then
            web_Curl = web_Curl & " --proxy-user " & Me.ProxyUsername & ":" & Me.ProxyPassword
        End If
    End If
    
    ' Setup Basic authentication
    If Me.Username <> "" Then
        web_Curl = web_Curl & " --user " & Me.Username & ":" & Me.Password
        Request.AddHeader "Authorization", "Basic " & WebHelpers.Base64Encode(Me.Username & ":" & Me.Password)
    End If
    
    ' Setup authenticator
    web_BeforeExecute Request
    
    ' Set headers and cookies
    Dim web_KeyValue As Dictionary
    For Each web_KeyValue In Request.Headers
        web_Curl = web_Curl & " -H '" & web_KeyValue("Key") & ": " & web_KeyValue("Value") & "'"
    Next web_KeyValue
    
    Dim web_CookieString As String
    For Each web_KeyValue In Request.Cookies
        web_CookieString = web_CookieString & web_KeyValue("Key") & "=" & web_KeyValue("Value") & ";"
    Next web_KeyValue
    web_Curl = web_Curl & " --cookie '" & web_CookieString & "'"
    
    ' Add method, data, and url
    web_Curl = web_Curl & " -X " & WebHelpers.MethodToName(Request.Method)
    web_Curl = web_Curl & " -d '" & Request.Body & "'"
    web_Curl = web_Curl & " '" & Me.GetFullUrl(Request) & "'"
    
    ' Log request and return
    WebHelpers.LogRequest Me, Request
    PrepareCurlRequest = web_Curl
    Exit Function
    
web_ErrorHandling:

    Err.Raise 11013 + vbObjectError, "WebClient.PrepareCURLRequest", _
        "An error occurred while preparing cURL request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

''
' Get full url for the given WebRequest
'
' @param {WebRequest} Request
' @return {String}
' --------------------------------------------- '
Public Function GetFullUrl(Request As WebRequest) As String
    GetFullUrl = WebHelpers.JoinUrl(Me.BaseUrl, Request.FormattedResource)
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub web_BeforeExecute(web_Request As WebRequest)
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, web_Request
    End If
    
    ' Preparing request includes adding headers
    ' -> Needs to happen after BeforeExecute in case headers were changed
    web_Request.Prepare
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMs = web_DefaultTimeoutMs
    Me.EnableAutoProxy = True
End Sub
