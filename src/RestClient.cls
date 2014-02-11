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
' RestClient v2.1.2
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Interact with REST web services from Excel
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const DefaultTimeoutMS As Integer = 5000


' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public BaseUrl As String
Public Authenticator As IAuthenticator
Public TimeoutMS As Integer
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
    Request.StartTimeoutTimer Me.TimeoutMS
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

Private Sub HttpSetup(ByRef Http As Object, ByRef Request As RestRequest, Optional UseAsync As Boolean = False)
    RestHelpers.PrepareHttpRequest Http, Request, Me.TimeoutMS, UseAsync
    
    If Me.ProxyServer <> "" Then
        RestHelpers.PrepareProxyForHttpRequest Http, Me.ProxyServer, Me.ProxyUsername, Me.ProxyPassword, Me.ProxyBypassList
    End If
    
    ' Before execute and http open hooks for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Request
        Me.Authenticator.HttpOpen Http, Request, Me.BaseUrl, UseAsync
    Else
        ' Nothing hooked in so open http object
        Http.Open Request.MethodName(), Request.FullUrl(Me.BaseUrl), UseAsync
    End If
    
    RestHelpers.SetHeaders Http, Request
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMS = DefaultTimeoutMS
End Sub
