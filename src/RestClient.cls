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

#If Mac Then
Private Declare Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As Long
Private Declare Function pclose Lib "libc.dylib" (ByVal File As Long) As Long
Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal Items As Long, ByVal stream As Long) As Long
Private Declare Function feof Lib "libc.dylib" (ByVal File As Long) As Long
#End If

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
    Dim Response As RestResponse
    
    ' Prepare request
    Set Request = PrepareRequest(Request)

#If Mac Then
    On Error GoTo ErrorHandling
    
    Dim cURL As String
    cURL = Me.PrepareCURL(Request)
    
    Dim File As Long
    Dim Chunk As String
    Dim Read As Long
    Dim Result As String
    Dim ExitCode As Long
    File = popen(cURL, "r")
    
    If File = 0 Then
        Err.Raise vbObjectError + 1, "RestClient.Execute", "An unknown error occurred while attempting to execute the request"
    End If
    
    While feof(File) = 0
        Chunk = VBA.space$(50)
        Read = fread(Chunk, 1, Len(Chunk) - 1, File)
        If Read > 0 Then
            Chunk = VBA.Left$(Chunk, Read)
            Result = Result & Chunk
        End If
    Loop
    
    ExitCode = pclose(File)
    
    ' TODO Check ExitCode
    
    Set Response = RestHelpers.CreateResponseFromCURL(Result, Request.ResponseFormat)
    RestHelpers.LogResponse Response, Request
    
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
        Set Response = RestHelpers.CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
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

''
' Prepare Request for execution
'
' @param {RestRequest} Request
' @return {RestRequest}
' --------------------------------------------- '
Public Function PrepareRequest(Request As RestRequest) As RestRequest
    ' Clone Request to "freeze" at time of execution
    Set PrepareRequest = Request.Clone
    
    ' Add general headers to request
    PrepareRequest.AddHeader "User-Agent", UserAgent
    PrepareRequest.AddHeader "Content-Type", Request.ContentType
    PrepareRequest.AddHeader "Accept", Request.Accept
    
    If PrepareRequest.IncludeContentLength Then
        PrepareRequest.AddHeader "Content-Length", PrepareRequest.ContentLength
    Else
        If PrepareRequest.Headers.Exists("Content-Length") Then
            PrepareRequest.Headers.Remove "Content-Length"
        End If
    End If
End Function

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
    Http.Open Request.MethodName(), Request.FullUrl(Me.BaseUrl), Async
    
    ' Set headers on http request (after open)
    RestHelpers.SetHeadersForHttp Http, Request
    
    ' Log request and return
    RestHelpers.LogRequest Request
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
    
    ' Prepare request
    Set Request = PrepareRequest(Request)
    
    ' Setup proxy?
    
    ' Setup Basic authentication
    
    ' Setup authenticator
    BeforeExecute Request
    
    ' Set headers
    
    ' Log request and return
    RestHelpers.LogRequest Request
    PrepareCURL = cURL
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub BeforeExecute(Request As RestRequest)
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, Request
    End If
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
