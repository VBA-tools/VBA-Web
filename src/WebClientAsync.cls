VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebClientAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' WebClientAsync v4.0.21
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebClientAsync` is a an asynchronous wrapper for the WebClient class,
'  providing a more intuitive way of running asynchronous requests than using
'  WebAsyncWrapper, and delivering results and errors back through VBA Events.
'
' _Note_ Windows-only and requires reference to "Microsoft WinHTTP Services, version 5.1"
'
' Usage:
'
' ```VBA
' Dim WithEvents Client As New WebClientAsync
' Client.BaseUrl = "https://www.example.com/api/"
'
' Dim Auth As New HttpBasicAuthenticator
' Auth.Setup Username, Password
' Set Client.Authenticator = Auth
'
' Dim Request As New WebRequest
' Dim Response As WebResponse
' ' Setup WebRequest...
'
' Client.ExecuteAsync(Request, uniqueArgs)
' ' -> Uses Http Basic authentication and appends Request.Resource to BaseUrl
'
' Private Sub Client_AsyncResponse(Response as WebResponse, uniqueArgs)
' ' Process response here
' End Sub
'
' Private Sub Client_AsyncError(Response as WebResponse, uniqueArgs)
' ' Handle response here
' End Sub
' ```
'
' Errors:
' 11010 / 80042b02 / -2147210494 - cURL error in Execute
' 11011 / 80042b03 / -2147210493 - Error in Execute
' 11012 / 80042b04 / -2147210492 - Error preparing http request
' 11080 - Active async requests must have unique RequestRef
' 11081 - Unable to find request to abort
'
' @class WebClientAsync
' @author Paul Freeman <emailtranslate@protopia.co.uk>
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit
' This is a Windows only module - so leave as blank class if Mac
#If Not Mac Then

' ============================================= '
' Global Class Declarations
' ============================================= '

' --------------------------------------------- '
' Module constants
' --------------------------------------------- '
Private Const DefaultMaxAsyncRequests As Integer = 8

' --------------------------------------------- '
' Windows Kernel timing functions
' --------------------------------------------- '

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function getFrequency Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function getFrequency Lib "kernel32" () As Long
#End If

' --------------------------------------------- '
' Public Class Variables
' --------------------------------------------- '

''
' Request that is currently executing.
'
' @property Request
' @type WebRequest
''
Public Request As WebRequest
Public MaxAsyncRequests As Integer

' --------------------------------------------- '
' Events created
' --------------------------------------------- '

Event AsyncResponse(ByVal Response As WebResponse, ByVal RequestRef As Variant)
Event AsyncError(ByVal ErrorNumber As Long, ByVal errorDescription As String, ByVal RequestRef As Variant)

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Client As WebClient
Private ActiveRequests As Collection
Private QueuedRequests As Collection
Private WithEvents WCAI As WebClientAsyncInstance
Attribute WCAI.VB_VarHelpID = -1
Private newWCAI As WebClientAsyncInstance
' ============================================= '
' Code
' ============================================= '
' --------------------------------------------- '
' Class Initialisation / Termination
' --------------------------------------------- '
Private Sub Class_Initialize()
    WebHelpers.LogDebug "Initializing", "WebClientAsync.Class_Initialize"
    
    Set ActiveRequests = New Collection
    
    Set QueuedRequests = New Collection
    Set Client = New WebClient
    ' WCAI is a Delegate object which is called to raise events
    Set WCAI = New WebClientAsyncInstance
    ' We create a newWCAI instance in advance in order to
    ' avoid delay to instantiate the class before issuing the web request
    Set newWCAI = New WebClientAsyncInstance
    Set newWCAI.Delegate = WCAI
    MaxAsyncRequests = DefaultMaxAsyncRequests
End Sub
Private Sub Class_Terminate()
    WebHelpers.LogDebug "Terminating", "WebClientAsync.Class_Terminate"
    
    Dim instance As WebClientAsyncInstance
    For Each instance In ActiveRequests
        WebHelpers.LogDebug "Aborting WebClientAsyncInstance: " & instance.nonce, "WebClientAsync.Abort"
        instance.Abort "WebClientAsync terminating"
        ActiveRequests.Remove instance.nonce
    Next instance
    Set ActiveRequests = Nothing
    Set Client = Nothing
End Sub

' --------------------------------------------- '
' Public Methods - Extending WebClient
' --------------------------------------------- '
Public Sub ExecuteAsync( _
        ByVal Request As WebRequest, _
        Optional ByVal RequestRef As Variant, _
        Optional ByVal Priority As Boolean _
        )
    WebHelpers.LogDebug "Called", "WebClientAsync.ExecuteAsync"
    
    Dim errMsg As String
    Dim nonce As String
    Dim instance As WebClientAsyncInstance
    
    ' RequestRef must be unique
    For Each instance In ActiveRequests
        If RequestRef = instance.RequestRef Then GoTo DuplicateReqArgs
    Next instance
    For Each instance In QueuedRequests
        If RequestRef = instance.RequestRef Then GoTo DuplicateReqArgs
    Next instance
    
    nonce = WebHelpers.CreateNonce
    newWCAI.nonce = nonce
    Set newWCAI.Request = Request.Clone
    Set newWCAI.Client = Client
    If IsObject(RequestRef) Then
        Set newWCAI.RequestRef = RequestRef
    Else
        newWCAI.RequestRef = RequestRef
    End If
    If Priority Then
        QueuedRequests.Add Item:=newWCAI, Key:=nonce, Before:=1
    Else
        QueuedRequests.Add Item:=newWCAI, Key:=nonce
    End If
    Call ExecuteQueuedRequests
    Set newWCAI = New WebClientAsyncInstance
    Set newWCAI.Delegate = WCAI
    Exit Sub

DuplicateReqArgs:
    Dim errorDescription As String

    errorDescription = "Active async requests must have unique RequestRef"
    WebHelpers.LogError errorDescription, "WebClientAsync.ExecuteAsync", 11080
    Err.Raise 11080 + vbObjectError, "WebClientAsync.ExecuteAsync", errorDescription

End Sub
Public Sub Abort(Optional ByVal Reason As String = "", Optional ByVal RequestRef As Variant)
    WebHelpers.LogDebug "Called", "WebClientAsync.Abort"
    
    Dim instance As WebClientAsyncInstance
    For Each instance In ActiveRequests
        If instance.RequestRef = RequestRef Then
            WebHelpers.LogWarning "Aborting active request: " & instance.nonce, "WebClientAsync.Abort"
            instance.Abort Reason
            ActiveRequests.Remove instance.nonce
            Exit Sub
        End If
    Next instance
    For Each instance In QueuedRequests
        If instance.RequestRef = RequestRef Then
            WebHelpers.LogWarning "Cancelling queued request: " & instance.nonce, "WebClientAsync.Abort"
            QueuedRequests.Remove instance.nonce
            Exit Sub
        End If
    Next instance
    
    Dim errorDescription As String
    errorDescription = "Unable to find request to abort"
    WebHelpers.LogError errorDescription, "WebClientAsync.Abort", 11081
    Err.Raise 11081 + vbObjectError, "WebClientAsync.ExecuteAsync", errorDescription
    
End Sub

Public Function Exists(Optional ByVal RequestRef As Variant) As Boolean
    For Each instance In ActiveRequests
        If instance.RequestRef = RequestRef Then
            Exists = True
            Exit Function
        End If
    Next instance
    For Each instance In QueuedRequests
        If instance.RequestRef = RequestRef Then
            Exists = True
            Exit Function
        End If
    Next instance
    Exists = False
End Function

' --------------------------------------------- '
' Event Handlers
' --------------------------------------------- '
Private Sub WCAI_AsyncError(ByVal nonce As String, ByVal ErrorNumber As Long, ByVal errorDescription As String, ByVal RequestRef As Variant)
    WebHelpers.LogDebug "Called", "WebClientAsync.AsyncError"
    ActiveRequests.Remove nonce
    RaiseEvent AsyncError(ErrorNumber, errorDescription, RequestRef)
    Call ExecuteQueuedRequests
End Sub

Private Sub WCAI_AsyncResponse(ByVal nonce As String, ByVal Response As WebResponse, ByVal RequestRef As Variant)
    WebHelpers.LogDebug "Called", "WebClientAsync.AsyncResponse"
    ActiveRequests.Remove nonce
    RaiseEvent AsyncResponse(Response, RequestRef)
End Sub

' --------------------------------------------- '
' Private Subs and Functions
' --------------------------------------------- '
Private Sub ExecuteQueuedRequests()
    Dim instance As WebClientAsyncInstance
    Do While ActiveRequests.Count < MaxAsyncRequests And QueuedRequests.Count > 0
        Set instance = QueuedRequests(1)
        QueuedRequests.Remove instance.nonce
        ActiveRequests.Add Item:=instance, Key:=instance.nonce
        
        On Error GoTo RequestError
        newWCAI.ExecuteAsyncInstance
        On Error GoTo 0
    Loop
    Exit Sub

RequestError:
    WebHelpers.LogError "WebClientAsyncInstance error " & Err.Number & ": " & Err.Description, "WebClientAsync.ExecuteAsync"
    ActiveRequests.Remove instance.nonce
    ' Rethrow error
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

' --------------------------------------------- '
' Public Properties - Passthrough to WebClient
' --------------------------------------------- '

Public Property Get BaseUrl() As String
    BaseUrl = Client.BaseUrl
End Property
Public Property Let BaseUrl(Value As String)
    Client.BaseUrl = Value
End Property

Public Property Get Authenticator() As IWebAuthenticator
    Set Authenticator = Client.Authenticator
End Property
Public Property Set Authenticator(Value As IWebAuthenticator)
    Set Client.Authenticator = Value
End Property

Public Property Get TimeoutMs() As Long
    TimeoutMs = Client.TimeoutMs
End Property
Public Property Let TimeoutMs(Value As Long)
    Client.TimeoutMs = Value
End Property

Public Property Get ProxyServer() As String
    ProxyServer = Client.ProxyServer
End Property
Public Property Let ProxyServer(Value As String)
    Client.ProxyServer = Value
End Property

Public Property Get ProxyBypassList() As String
    ProxyBypassList = Client.ProxyBypassList
End Property
Public Property Let ProxyBypassList(Value As String)
    Client.ProxyBypassList = Value
End Property

Public Property Get ProxyUsername() As String
    ProxyUsername = Client.ProxyUsername
End Property
Public Property Let ProxyUsername(Value As String)
    Client.ProxyUsername = Value
End Property

Public Property Get ProxyPassword() As String
    ProxyPassword = Client.ProxyPassword
End Property
Public Property Let ProxyPassword(Value As String)
    Client.ProxyPassword = Value
End Property

Public Property Get EnableAutoProxy() As Boolean
    EnableAutoProxy = Client.EnableAutoProxy
End Property
Public Property Let EnableAutoProxy(Value As Boolean)
    Client.EnableAutoProxy = Value
End Property

Public Property Get Insecure() As Boolean
    Insecure = Client.Insecure
End Property
Public Property Let Insecure(Value As Boolean)
    Client.Insecure = Value
End Property

Public Property Get FollowRedirects() As Boolean
    FollowRedirects = Client.FollowRedirects
End Property
Public Property Let FollowRedirects(Value As Boolean)
    Client.FollowRedirects = Value
End Property

' --------------------------------------------- '
' Public Methods - Passthrough to WebClient
' --------------------------------------------- '

Public Function Execute(Request As WebRequest) As WebResponse
    ' Although this is intended as an async client,
    ' the sync execution functions are mapped in order to
    ' allow the user to choose sync/async at run time.
    WebHelpers.LogDebug "WebAsyncRequest.Execute called", "WebClientAsync.Execute"
    Execute = Client.Execute(Request)
End Function

Public Function GetJson(Url As String, Optional Options As Dictionary = Nothing) As WebResponse
    GetJson = Client.GetJson(Url, Options)
End Function

Public Function PostJson(Url As String, Body As Variant, Optional Options As Dictionary = Nothing) As WebResponse
    PostJson = Client.PostJson(Url, Body, Options)
End Function

Public Function GetFullUrl(Request As WebRequest) As String
    GetFullUrl = Client.GetFullUrl(Request)
End Function

Public Sub SetProxy( _
        Server As String, _
        Optional Username As String = "", _
        Optional Password As String = "", _
        Optional BypassList As String = "")
    Client.SetProxy Server, Username, Password, BypassList
End Sub

#End If

