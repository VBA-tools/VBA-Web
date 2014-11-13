VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' RestClient v4.0.0-beta.2
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Interact with REST web services from Excel
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const UserAgent As String = "Excel Client v4.0.0-beta.2 (https://github.com/timhall/Excel-REST)"
Private Const DefaultTimeoutMS As Long = 5000

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public BaseUrl As String
Public Authenticator As IAuthenticator
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
' @param {RestRequest} request The request to execute
' @return {RestResponse} Wrapper of server response for request
' --------------------------------------------- '

Public Function Execute(Request As RestRequest) As RestResponse
    Dim Response As RestResponse

#If Mac Then
    On Error GoTo ErrorHandling
    
    Dim cURL As String
    Dim Result As ShellResult
    
    cURL = Me.PrepareCURL(Request)
    Result = RestHelpers.ExecuteInShell(cURL)
    
    Set Response = RestHelpers.CreateResponseFromCURL(Result, Request.ResponseFormat)
    RestHelpers.LogResponse Response, Request
    Set Execute = Response
    Exit Function
    
ErrorHandling:

    ' Rethrow error
    RestHelpers.LogError Err.Description, "RestClient.Execute", Err.Number
    Err.Raise Err.Number, Source:=Err.Source, Description:=Err.Description
#Else
    On Error GoTo ErrorHandling
    
    Dim Http As Object
    Set Http = Me.PrepareHttpRequest(Request)
    
    Http.Send Request.Body
    Do While Not Http.WaitForResponse(25)
        DoEvents
    Loop
    
    Set Response = RestHelpers.CreateResponseFromHttp(Http, Request.ResponseFormat)
    RestHelpers.LogResponse Response, Request
    
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
        Set Response = RestHelpers.CreateResponse(WebStatusCode.RequestTimeout, "Request Timeout")
        RestHelpers.LogResponse Response, Request
        Set Execute = Response
        Err.Clear
    Else
        ' Rethrow error
        RestHelpers.LogError Err.Description, "RestClient.Execute", Err.Number
        Err.Raise Err.Number, Source:=Err.Source, Description:=Err.Description
    End If
#End If
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
' @return {RestResponse} Response
' --------------------------------------------- '

Public Function GetJSON(Url As String, Optional Options As Dictionary) As RestResponse
    Dim Request As RestRequest
    Set Request = RestHelpers.CreateRequestFromOptions(Options)
    Request.Resource = Url
    Request.Format = WebFormat.json
    Request.Method = WebMethod.httpGET
    
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
' @return {RestResponse} Response
' --------------------------------------------- '

Public Function PostJSON(Url As String, Body As Variant, Optional Options As Dictionary) As RestResponse
    Dim Request As RestRequest
    Set Request = RestHelpers.CreateRequestFromOptions(Options)
    Request.Resource = Url
    Request.Format = WebFormat.json
    Request.Method = WebMethod.httpPOST
    Request.Body = Body
    
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
' Prepare Http request for given RestRequest
'
' @param {RestRequest} Request
' @return {WinHttpRequest}
' --------------------------------------------- '

Public Function PrepareHttpRequest(Request As RestRequest, Optional Async As Boolean = True) As Object
    Dim Http As Object
    Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Set timeouts
    Http.SetTimeouts Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS
    
    ' Setup proxy
    If Me.ProxyServer <> "" Then
        LogDebug "SetProxy: " & Me.ProxyServer, "RestClient.PrepareHttpRequest"
        Http.SetProxy 2, Me.ProxyServer, Me.ProxyBypassList
        
        If Me.ProxyUsername <> "" Then
            LogDebug "SetProxyCredentials: " & Me.ProxyUsername & ", " & RestHelpers.Obfuscate(Me.ProxyPassword), "RestClient.PrepareHttpRequest"
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
        Request.AddHeader "Authorization", "Basic " & RestHelpers.Base64Encode(Me.Username & ":" & Me.Password)
    End If
    
    ' Setup authenticator
    BeforeExecute Request
    
    ' Open http request
    Http.Open RestHelpers.MethodToName(Request.Method), Me.GetFullUrl(Request.FormattedResource), Async
    
    ' Set headers on http request (after open)
    RestHelpers.SetHeadersForHttp Http, Request
    
    ' Log request and return
    RestHelpers.LogRequest Request, Me.GetFullUrl(Request.FormattedResource)
    Set PrepareHttpRequest = Http
End Function

''
' Prepare cURL command for given RestRequest
'
' @param {RestRequest} Request
' @return {String}
' --------------------------------------------- '

Public Function PrepareCURL(Request As RestRequest) As String
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
        Request.AddHeader "Authorization", "Basic " & RestHelpers.Base64Encode(Me.Username & ":" & Me.Password)
    End If
    
    ' Setup authenticator
    BeforeExecute Request
    
    ' Set headers and cookies
    Dim HeaderKey As Variant
    For Each HeaderKey In Request.Headers.Keys
        cURL = cURL & " -H '" & HeaderKey & ": " & Request.Headers(HeaderKey) & "'"
    Next HeaderKey
    
    Dim CookieKey As Variant
    Dim CookieString As String
    For Each CookieKey In Request.Cookies.Keys
        CookieString = CookieString & CookieKey & "=" & Request.Cookies(CookieKey) & ";"
    Next CookieKey
    cURL = cURL & " --cookie '" & CookieString & "'"
    
    ' Add method, data, and url
    cURL = cURL & " -X " & RestHelpers.MethodToName(Request.Method)
    cURL = cURL & " -d '" & Request.Body & "'"
    cURL = cURL & " '" & Me.GetFullUrl(Request.FormattedResource) & "'"
    
    ' Log request and return
    RestHelpers.LogRequest Request, Me.GetFullUrl(Request.FormattedResource)
    PrepareCURL = cURL
End Function

''
' Get full url given the request resource
'
' @param {String} Resource
' @return {String}
' --------------------------------------------- '

Public Function GetFullUrl(Resource As String) As String
    GetFullUrl = RestHelpers.JoinUrl(Me.BaseUrl, Resource)
End Function

' ============================================= '
' Private Methods
' ============================================= '

''
' Prepare Request for execution
'
' @param {RestRequest} Request
' --------------------------------------------- '
Private Sub PrepareRequest(ByRef Request As RestRequest)
    ' Add general headers to request
    Request.AddHeader "User-Agent", UserAgent
    Request.AddHeader "Content-Type", Request.ContentType
    Request.AddHeader "Accept", Request.Accept
    Request.AddHeader "Content-Length", CStr(Request.ContentLength)
End Sub

Private Sub BeforeExecute(Request As RestRequest)
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, Request
    End If
    
    ' Prepare request (after BeforeExecute)
    PrepareRequest Request
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
