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
' WebClient v4.0.0-beta.3
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Interact with REST web services from Excel
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const DefaultTimeoutMS As Long = 5000

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public BaseUrl As String
Public Authenticator As IWebAuthenticator
Public TimeoutMS As Long
Public Username As String
Public Password As String
Public ProxyServer As String
Public ProxyUsername As String
Public ProxyPassword As String
Public ProxyBypassList As String

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
    Dim Http As Object
    Dim Response As New WebResponse
    
    On Error GoTo ErrorHandling

#If Mac Then
    Dim cURL As String
    Dim Result As WebShellResult
    
    cURL = Me.PrepareCURLRequest(Request)
    Result = WebHelpers.ExecuteInShell(cURL)
    
    If Result.ExitCode > 0 Then
        Dim ErrorNumber As Long
        Dim ErrorMessage As String
        ErrorNumber = Result.ExitCode / 256
        
        Select Case ErrorNumber
        Case 5, 6, 7, 12, 28
            ' 5 = CURLE_COULDNT_RESOLVE_PROXY
            ' 6 = CURLE_COULDNT_RESOLVE_HOST
            ' 7 = CURLE_COULDNT_CONNECT
            ' 12 = CURLE_FTP_ACCEPT_TIMEOUT
            ' 28 = CURLE_OPERATION_TIMEDOUT
            
            Err.Raise 1 Or vbObjectError, "WebClient.Execute", "The operation timed out"
        Case 1
            ' 1 = CURLE_UNSUPPORTED_PROTOCOL
            Err.Raise vbObjectError + 1, "The protocol used in the request is not supported (1: CURLE_UNSUPPORTED_PROTOCOL)" & vbNewLine & _
                "URL: " & Me.GetFullRequestUrl(Request) & vbNewLine & _
                "Protocol: " & WebHelpers.UrlParts(Me.GetFullRequestUrl(Request))("Protocol")
        Case 3
            ' 3 = CURLE_URL_MALFORMAT
            Err.Raise vbObjectError + 1, "The URL was not properly formatted (3: CURLE_URL_MALFORMAT)" & _
                "URL: " & Me.GetFullRequestUrl(Request)
        Case Else
            Err.Raise vbObjectError + 1, "cURL Error #" & ErrorNumber & vbNewLine & _
                "Details: http://curl.haxx.se/libcurl/c/libcurl-errors.html"
        End Select
    End If
    
    Response.CreateFromCURL Me, Request, Result.Output

#Else
    Set Http = Me.PrepareHttpRequest(Request)
    
    Http.Send Request.Body
    Do While Not Http.WaitForResponse(25)
        DoEvents
    Loop
    
    Response.CreateFromHttp Me, Request, Http
    
#End If

    WebHelpers.LogResponse Me, Request, Response
    
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.AfterExecute Me, Request, Response
    End If
    
    Set Http = Nothing
    Set Execute = Response
    Exit Function
    
ErrorHandling:

    Set Http = Nothing
    If InStr(Err.Description, "The operation timed out") > 0 Or InStr(Err.Description, "A connection with the server could not be established") > 0 Then
        ' Return 408
        Response.StatusCode = WebStatusCode.RequestTimeout
        Response.StatusDescription = "Request Timeout"
        
        WebHelpers.LogResponse Me, Request, Response
        Set Execute = Response
        Err.Clear
    Else
        ' Rethrow error
        WebHelpers.LogError Err.Description, "WebClient.Execute", Err.Number
        Err.Raise Err.Number, Source:=Err.Source, Description:=Err.Description
    End If
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

Public Function GetJSON(Url As String, Optional Options As Dictionary) As WebResponse
    Dim Request As New WebRequest
    Request.CreateFromOptions Options
    Request.Resource = Url
    Request.Format = WebFormat.Json
    Request.Method = WebMethod.HttpGet
    
    Set GetJSON = Me.Execute(Request)
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

Public Function PostJSON(Url As String, Body As Variant, Optional Options As Dictionary) As WebResponse
    Dim Request As New WebRequest
    Request.CreateFromOptions Options
    Request.Resource = Url
    Request.Format = WebFormat.Json
    Request.Method = WebMethod.HttpPost
    If IsObject(Body) Then
        Set Request.Body = Body
    Else
        Request.Body = Body
    End If
    
    Set PostJSON = Me.Execute(Request)
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
' @param {WebRequest} Request
' @return {WinHttpRequest}
' --------------------------------------------- '

Public Function PrepareHttpRequest(Request As WebRequest, Optional Async As Boolean = True) As Object
    Dim Http As Object
    Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Set timeouts
    Http.SetTimeouts Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS
    
    ' Setup authenticator
    BeforeExecute Request
    
    ' Open http request
    Http.Open WebHelpers.MethodToName(Request.Method), Me.GetFullRequestUrl(Request), Async
    
    ' Setup proxy
    If Me.ProxyServer <> "" Then
        LogDebug "SetProxy: " & Me.ProxyServer, "WebClient.PrepareHttpRequest"
        Http.SetProxy 2, Me.ProxyServer, Me.ProxyBypassList
        
        If Me.ProxyUsername <> "" Then
            LogDebug "SetProxyCredentials: " & Me.ProxyUsername & ", " & WebHelpers.Obfuscate(Me.ProxyPassword), "WebClient.PrepareHttpRequest"
            Http.SetCredentials Me.ProxyUsername, Me.ProxyPassword, 1 ' 1 = Credentials for proxy
        End If
    Else
        ' Attempt to get proxy setup with Proxycfg.exe, otherwise direct
        ' See http://msdn.microsoft.com/en-us/library/windows/desktop/aa384059(v=vs.85).aspx for details
        Http.SetProxy 1 ' 1 = HTTPREQUEST_PROXYSETTING_PRECONFIG
    End If
    
    ' Setup Basic authentication
    If Me.Username <> "" Then
        Http.SetCredentials Me.Username, Me.Password, 0 ' 0 = Credentials for server
        Request.AddHeader "Authorization", "Basic " & WebHelpers.Base64Encode(Me.Username & ":" & Me.Password)
    End If
    
    ' Set headers on http request (after open)
    Dim KeyValue As Dictionary
    For Each KeyValue In Request.Headers
        Http.setRequestHeader KeyValue("Key"), KeyValue("Value")
    Next KeyValue
    
    For Each KeyValue In Request.Cookies
        Http.setRequestHeader "Cookie", KeyValue("Key") & "=" & KeyValue("Value")
    Next KeyValue
    
    ' Log request and return
    WebHelpers.LogRequest Me, Request
    Set PrepareHttpRequest = Http
End Function

''
' Prepare cURL request for given WebRequest
'
' @param {WebRequest} Request
' @return {String}
' --------------------------------------------- '

Public Function PrepareCURLRequest(Request As WebRequest) As String
    Dim cURL As String
    cURL = "curl -i"
    
    ' Set timeouts
    cURL = cURL & " --connect-timeout " & Me.TimeoutMS / 1000
    cURL = cURL & " --max-time " & 3 * Me.TimeoutMS / 1000
    
    ' Setup proxy
    If Me.ProxyServer <> "" Then
        cURL = cURL & " --proxy " & Me.ProxyServer
        
        If Me.ProxyBypassList <> "" Then
            cURL = cURL & " --noproxy " & Me.ProxyBypassList
        End If
        If Me.ProxyUsername <> "" Then
            cURL = cURL & " --proxy-user " & Me.ProxyUsername & ":" & Me.ProxyPassword
        End If
    End If
    
    ' Setup Basic authentication
    If Me.Username <> "" Then
        cURL = cURL & " --user " & Me.Username & ":" & Me.Password
        Request.AddHeader "Authorization", "Basic " & WebHelpers.Base64Encode(Me.Username & ":" & Me.Password)
    End If
    
    ' Setup authenticator
    BeforeExecute Request
    
    ' Set headers and cookies
    Dim KeyValue As Dictionary
    For Each KeyValue In Request.Headers
        cURL = cURL & " -H '" & KeyValue("Key") & ": " & KeyValue("Value") & "'"
    Next KeyValue
    
    Dim CookieString As String
    For Each KeyValue In Request.Cookies
        CookieString = CookieString & KeyValue("Key") & "=" & KeyValue("Value") & ";"
    Next KeyValue
    cURL = cURL & " --cookie '" & CookieString & "'"
    
    ' Add method, data, and url
    cURL = cURL & " -X " & WebHelpers.MethodToName(Request.Method)
    cURL = cURL & " -d '" & Request.Body & "'"
    cURL = cURL & " '" & Me.GetFullRequestUrl(Request) & "'"
    
    ' Log request and return
    WebHelpers.LogRequest Me, Request
    PrepareCURLRequest = cURL
End Function

''
' Get full url for the given path
'
' @param {String} Path
' @return {String}
' --------------------------------------------- '

Public Function GetFullUrl(Path As String) As String
    GetFullUrl = WebHelpers.JoinUrl(Me.BaseUrl, Path)
End Function

''
' Get full url for the given WebRequest
'
' @param {WebRequest} Request
' @return {String}
' --------------------------------------------- '
Public Function GetFullRequestUrl(Request As WebRequest) As String
    GetFullRequestUrl = GetFullUrl(Request.FormattedResource)
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub BeforeExecute(Request As WebRequest)
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, Request
    End If
    
    ' Preparing request includes adding headers
    ' -> Needs to happen after BeforeExecute in case headers were changed
    Request.Prepare
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
