Attribute VB_Name = "RestClientBase"
''
' RestClientBase v2.1.0
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

' > Customize with BaseUrl and other properties
Private Sub Initialize()
    ' BaseUrl = "https://..."
    ' ProxyServer = "..."
    ' ProxyUsername = "..."
    ' ProxyPassword = "..."
    
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
    Dim Response As RestResponse
    Dim Http As Object
    Dim HeaderKey As Variant
    
    On Error GoTo ErrorHandling
    Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HttpSetup Http, Request, False
    
    ' Send the request
    Http.send Request.Body
    
    ' Handle response
    Set Response = Request.CreateResponseFromHttp(Http)
    
ErrorHandling:

    If Not Http Is Nothing Then Set Http = Nothing
    
    If Err.Number <> 0 Then
        If InStr(Err.Description, "The operation timed out") > 0 Then
            ' Return 408
            Set Response = Request.CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
            Err.Clear
        Else
            ' Rethrow error
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
    Dim Response As New RestResponse
    Dim Http As Object
    
    On Error GoTo ErrorHandling
    
    ' Setup the request
    Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HttpSetup Http, Request, True
    Request.Callback = Callback
    Request.CallbackArgs = CallbackArgs
    
    ' Send the request
    Request.StartTimeoutTimer TimeoutMS
    Http.send Request.Body
    
    ' Clean up and return
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    If Not Http Is Nothing Then Set Http = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Description:=Err.Description
    End If
End Function

Private Sub HttpSetup(ByRef Http As Object, ByRef Request As RestRequest, Optional UseAsync As Boolean = False)
    If Not Initialized Then: Initialize
    
    RestHelpers.PrepareHttpRequest Http, Request, TimeoutMS, UseAsync
    
    
    If ProxyServer <> "" Then
        RestHelpers.PrepareProxyForHttpRequest Http, ProxyServer, ProxyUsername, ProxyPassword, ProxyBypassList
    End If
    
    ' Before execute and http open hooks for authentication
    BeforeExecute Request
    HttpOpen Http, Request, BaseUrl, UseAsync
    
    RestHelpers.SetHeaders Http, Request
End Sub

