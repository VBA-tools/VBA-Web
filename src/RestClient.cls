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
' RestClient v3.1.4
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Interact with REST web services from Excel
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const UserAgent As String = "Excel Client v3.1.4 (https://github.com/timhall/Excel-REST)"
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
Public ProxyBypassList As Variant

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
    Set Execute = ExecuteRequest(Request)
End Function

''
' Execute the specified request asynchronously
'
' @param {RestRequest} Request The request to execute
' @param {String} Callback Name of function to call when request completes
' @param {Variant} [CallbackArgs] Variable array of arguments that get passed directly to callback function
' @return {Boolean} Status of initiating request
' --------------------------------------------- '

Public Function ExecuteAsync(Request As RestRequest, Optional Callback As String = "", Optional ByVal CallbackArgs As Variant) As Boolean
    ExecuteRequest Request, Callback, CallbackArgs
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
    Request.Format = AvailableFormats.json
    Request.Method = AvailableMethods.httpGET
    
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
    Request.Format = AvailableFormats.json
    Request.Method = AvailableMethods.httpPOST
    Request.AddBody Body
    
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
    Optional Username As String = "", Optional Password As String = "", Optional BypassList As Variant)
    
    Me.ProxyServer = ProxyServer
    Me.ProxyUsername = Username
    Me.ProxyPassword = Password
    Me.ProxyBypassList = BypassList
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Function ExecuteRequest(Request As RestRequest, Optional Callback As String = "", Optional ByVal CallbackArgs As Variant) As RestResponse
    Dim IsAsync As Boolean
    If Callback <> "" Then: IsAsync = True
    
    ' Clone Request to "freeze" at time of execution and log
    Set Request = Request.Clone
    RestHelpers.LogRequest Request
    
#If Mac Then
    On Error GoTo ErrorHandling
    
    ' TODO
    
ErrorHandling:
    
#Else
    On Error GoTo ErrorHandling
    Dim Http As New WinHttpRequest
    
    ' Set timeouts
    Http.setTimeouts Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS
    
    ' Setup proxy
    If Me.ProxyServer <> "" Then
        LogDebug "SetProxy: " & Me.ProxyServer, "RestHelpers.PrepareProxyForHttpRequest"
        Http.SetProxy 2, Me.ProxyServer, Me.ProxyBypassList
        
        If Me.ProxyUsername <> "" Then
            LogDebug "SetProxyCredentials: " & Me.ProxyUsername & ", " & Obfuscate(Me.ProxyPassword), "RestClient.ExecuteRequest"
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
    
    ' Before execute hook for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, Request
        
        Set Request.Authenticator = Me.Authenticator
        Set Request.Client = Me
    End If
    
    ' Open http request
    Http.Open Request.MethodName(), Request.FullUrl(Me.BaseUrl), True
    
    ' Add general headers to request
    Request.AddHeader "User-Agent", UserAgent
    Request.AddHeader "Content-Type", Request.ContentType
    Request.AddHeader "Accept", Request.Accept
    
    If Request.IncludeContentLength Then
        Request.AddHeader "Content-Length", Request.ContentLength
    Else
        If Request.Headers.Exists("Content-Length") Then
            Request.Headers.Remove "Content-Length"
        End If
    End If
    
    ' Set headers on http request (after open)
    RestHelpers.SetHeadersForHttp Http, Request
    
    If IsAsync Then
        ' Pass http, callback, and args to request for async
        Set Request.HttpRequest = Http
        Request.Callback = Callback
        Request.CallbackArgs = CallbackArgs
        
        Request.StartTimeoutTimer Me.TimeoutMS
    End If
    
    Http.Send Request.Body
    
    If Not IsAsync Then
        ' Wait for the response
        Do While Not Http.WaitForResponse(25)
            DoEvents
        Loop
        
        ' Handle the response
        Dim Response As RestResponse
        Set Response = RestHelpers.CreateResponseFromHttp(Http, Request.ResponseFormat)
        RestHelpers.LogResponse Response, Request
        
        If Not Me.Authenticator Is Nothing Then
            Me.Authenticator.AfterExecute Me, Request, Response
        End If
        
        Set Http = Nothing
        Set ExecuteRequest = Response
    End If
    
    Exit Function

ErrorHandling:
    Set Http = Nothing
    If InStr(Err.Description, "The operation timed out") > 0 Or InStr(Err.Description, "A connection with the server could not be established") > 0 Then
        ' Return 408
        Set Response = RestHelpers.CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
        RestHelpers.LogResponse Response, Request
        Set ExecuteRequest = Response
        Err.Clear
    Else
        ' Rethrow error
        RestHelpers.LogError Err.Description, "RestClient.Execute", Err.Number
        Err.Raise Err.Number, Description:=Err.Description
    End If
        
#End If
End Function

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
