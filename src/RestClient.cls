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

Private Const DefaultTimeoutMS As Long = 5000


' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public BaseUrl As String
Public Authenticator As IAuthenticator
Public TimeoutMS As Long
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
    On Error GoTo ErrorHandling
    Dim Http As MSXML2.ServerXMLHTTP60
    Dim Response As RestResponse
    
    ' Clone Request to "freeze" at time of execution
    Set Request = Request.Clone
    
    ' Setup http
    Set Http = SetupHttp(Request, False)
    
    ' Send the request
    RestHelpers.LogRequest Request
    Http.send Request.Body
    
    ' Handle the response
    Set Response = RestHelpers.CreateResponseFromHttp(Http, Request.ResponseFormat)
    RestHelpers.LogResponse Response, Request
    
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.AfterExecute Me, Request, Response
    End If
    
ErrorHandling:
    
    If Not Http Is Nothing Then Set Http = Nothing
    If Err.Number <> 0 Then
        If InStr(Err.Description, "The operation timed out") > 0 Then
            ' Return 408
            Set Response = RestHelpers.CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
            RestHelpers.LogResponse Response, Request
            Err.Clear
        Else
            ' Rethrow error
            RestHelpers.LogError Err.Description, "RestClient.Execute", Err.Number
            Err.Raise Err.Number, Description:=Err.Description
        End If
    End If
    
    Set Execute = Response
End Function

''
' Execute the specified request asynchronously
'
' @param {RestRequest} request The request to execute
' @param {String} callback Name of function to call when request completes (specify "" if none)
' @param {Variant} [callbackArgs] Variable array of arguments that get passed directly to callback function
' @return {Boolean} Status of initiating request
' --------------------------------------------- '

Public Function ExecuteAsync(Request As RestRequest, Callback As String, Optional ByVal CallbackArgs As Variant) As Boolean
    On Error GoTo ErrorHandling
    Dim Http As MSXML2.ServerXMLHTTP60
    
    ' Clone Request to "freeze" at time of execution
    Set Request = Request.Clone
    Request.Callback = Callback
    Request.CallbackArgs = CallbackArgs
    
    ' Setup http
    Set Http = SetupHttp(Request, UseAsync:=True)
    
    ' Send the request
    Request.StartTimeoutTimer Me.TimeoutMS
    RestHelpers.LogRequest Request
    Http.send Request.Body
    
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    ' Close Http and rethrow error
    If Not Http Is Nothing Then Set Http = Nothing
    RestHelpers.LogError Err.Description, "RestClient.ExecuteAsync", Err.Number
    Err.Raise Err.Number, Description:=Err.Description
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

Private Function SetupHttp(ByRef Request As RestRequest, Optional UseAsync As Boolean = False) As Object
    Dim Http As New MSXML2.ServerXMLHTTP60
    
    ' Set timeouts
    Http.setTimeouts Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS, Me.TimeoutMS
    
    ' Pass http to request and setup onreadystatechange
    If UseAsync Then
        Set Request.HttpRequest = Http
        Http.onreadystatechange = Request
    End If
    
    ' Setup proxy
    If ProxyServer <> "" Then
        LogDebug "SetProxy: " & ProxyServer, "RestHelpers.PrepareProxyForHttpRequest"
        Http.SetProxy 2, Me.ProxyServer, Me.ProxyBypassList
        
        If Me.ProxyUsername <> "" Then
            LogDebug "SetProxyCredentials: " & Me.ProxyUsername & ", " & Obfuscate(Me.ProxyPassword), "RestClient.SetupHttp"
            Http.SetProxyCredentials Me.ProxyUsername, Me.ProxyPassword
        End If
    End If
    
    ' Before execute and http open hooks for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, Request
        Me.Authenticator.HttpOpen Http, Me, Request, Me.BaseUrl, UseAsync
        Set Request.Authenticator = Me.Authenticator
        Set Request.Client = Me
    Else
        ' Nothing hooked in so open http object
        Http.Open Request.MethodName(), Request.FullUrl(Me.BaseUrl), UseAsync
    End If
    
    ' Set headers on http request (after open)
    RestHelpers.SetHeaders Http, Request
    
    Set SetupHttp = Http
End Function

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
