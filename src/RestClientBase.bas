Attribute VB_Name = "RestClientBase"
''
' RestClientBase v2.2.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Extendable RestClientBase for developing custom client classes
' - Embed authenticator logic with BeforeExecute and HttpOpen methods
' - Add public methods and helpers for specific requests
'
' Look for ">" for points to customize
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' > Customize with BaseUrl
Private Const DefaultBaseUrl As String = ""
Private Const TimeoutMS As Integer = 5000
Private Initialized As Boolean

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public BaseUrl As String
Public ProxyServer As String
Public ProxyUsername As String
Public ProxyPassword As String
Public ProxyBypassList As Variant

' ============================================= '
' Public Methods
' ============================================= '

' > Customize with public methods and helpers here

' ============================================= '
' Extend RestClientBase
' ============================================= '

' > Customize to update request before execution (matches IAuthenticator)
Private Sub BeforeExecute(ByRef Request As RestRequest)
    
End Sub

' > Customize to perform special http open behavior (matches IAuthenticator)
Private Sub HttpOpen(ByRef Http As Object, ByRef Request As RestRequest, BaseUrl As String, Optional UseAsync As Boolean = False)
    Http.Open Request.MethodName(), Request.FullUrl(BaseUrl), UseAsync
End Sub


Private Sub Initialize()
    If BaseUrl = "" Then
        BaseUrl = DefaultBaseUrl
    End If
    
    ' > Customize with any properties
    
    Initialized = True
End Sub

' ============================================= '
' Internal Methods
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
    
    Set Http = HttpSetup(Request, False)
    Set Execute = RestHelpers.ExecuteRequest(Http, Request)
    
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
    Set Http = HttpSetup(Request, True)
    RestHelpers.ExecuteRequestAsync Http, Request, TimeoutMS, Callback, CallbackArgs
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    ' Close Http and rethrow error
    If Not Http Is Nothing Then Set Http = Nothing
    Err.Raise Err.Number, Description:=Err.Description
End Function

''
' Setup proxy server
'
' @param {String} ProxyServer
' @param {String} [Username=""]
' @param {String} [Password=""]
' @param {Variant} [BypassList]
' --------------------------------------------- '

Public Sub SetupProxy(ProxyServer As String, _
    Optional Username As String = "", Optional Password As String = "", Optional BypassList As Variant)
    
    ProxyServer = ProxyServer
    ProxyUsername = ProxyUsername
    ProxyPassword = ProxyPassword
    BypassList = BypassList
End Sub

Private Function HttpSetup(ByRef Request As RestRequest, Optional UseAsync As Boolean = False) As Object
    If Not Initialized Then: Initialize
    
    Set HttpSetup = RestHelpers.PrepareHttpRequest(Request, TimeoutMS, UseAsync)

    If ProxyServer <> "" Then
        RestHelpers.PrepareProxyForHttpRequest HttpSetup, ProxyServer, ProxyUsername, ProxyPassword, ProxyBypassList
    End If
    
    ' Before execute and http open hooks for authentication
    BeforeExecute Request
    HttpOpen HttpSetup, Request, BaseUrl, UseAsync
    
    RestHelpers.SetHeaders HttpSetup, Request
End Function
