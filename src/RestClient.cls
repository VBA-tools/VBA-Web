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
' RestClient v3.0.6
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
    Dim Http As Object
    
    RestHelpers.LogDebug Request.FullUrl(Me.BaseUrl), "RestClient.Execute"
    Set Http = HttpSetup(Request, False)
    Set Execute = RestHelpers.ExecuteRequest(Http, Request)
    
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.AfterExecute Me, Request, Execute
    End If
    
ErrorHandling:

    If Not Http Is Nothing Then Set Http = Nothing
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Description:=Err.Description
    End If
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
    Dim Http As Object
    
    ' Setup the request
    RestHelpers.LogDebug Request.FullUrl(Me.BaseUrl), "RestClient.ExecuteAsync"
    Set Http = HttpSetup(Request, True)
    RestHelpers.ExecuteRequestAsync Http, Request, Me.TimeoutMS, Callback, CallbackArgs
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    ' Close Http and rethrow error
    If Not Http Is Nothing Then Set Http = Nothing
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
    Request.Format = json
    Request.Method = httpGET
    
    RestHelpers.LogDebug Request.FullUrl(Me.BaseUrl), "RestClient.GetJSON"
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
    Request.Format = json
    Request.Method = httpPOST
    Request.AddBody Body
    
    RestHelpers.LogDebug Request.FullUrl(Me.BaseUrl), "RestClient.PostJSON"
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

Private Function HttpSetup(ByRef Request As RestRequest, Optional UseAsync As Boolean = False) As Object
    Set HttpSetup = RestHelpers.PrepareHttpRequest(Request, Me.TimeoutMS, UseAsync)
    
    If Me.ProxyServer <> "" Then
        RestHelpers.PrepareProxyForHttpRequest HttpSetup, Me.ProxyServer, Me.ProxyUsername, Me.ProxyPassword, Me.ProxyBypassList
    End If
    
    ' Before execute and http open hooks for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, Request
        Me.Authenticator.HttpOpen HttpSetup, Me, Request, Me.BaseUrl, UseAsync
        Set Request.Authenticator = Me.Authenticator
        Set Request.Client = Me
    Else
        ' Nothing hooked in so open http object
        HttpSetup.Open Request.MethodName(), Request.FullUrl(Me.BaseUrl), UseAsync
        RestHelpers.LogDebug "Http.Open " & Request.MethodName() & Request.FullUrl(Me.BaseUrl), "RestClient.HttpSetup"
    End If
    
    RestHelpers.SetHeaders HttpSetup, Request
End Function

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
