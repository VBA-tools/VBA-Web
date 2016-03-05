VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebClientAsyncInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' WebClientAsyncInstance v4.0.21
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebClientAsyncInstance` is an INTERNAL helper class for WebClientAsync.
'  A helper class is required to allow the same WebClientAsync instance
'  to be used for multiple simultaneous web calls. In essence, the HTTP
'  request sends results back through an event. The only way to identify
'  the specific call that produced the response is to have the HTTP instance
'  associated with a unique event processor i.e. a unique object instance.
'
' _Note_ Windows-only and requires reference to "Microsoft WinHTTP Services, version 5.1"
'
' @class WebClientAsyncInstance
' @author Paul Freeman <emailtranslate@protopia.co.uk>
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit
#If Not Mac Then

' --------------------------------------------- '
' Events created
' --------------------------------------------- '

Event AsyncResponse(ByVal nonce As String, ByVal Response As WebResponse, ByVal RequestRef As Variant)
Event AsyncError(ByVal nonce As String, ByVal ErrorNumber As Long, ByVal errorDescription As String, ByVal RequestRef As Variant)

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '
Private RequestStartTicks As Long
Private TimerTickFrequency As Currency

' --------------------------------------------- '
' Public Variables
' --------------------------------------------- '

Public nonce As String
Public Request As WebRequest
Public Client As WebClient
Public WithEvents Http As WinHttpRequest
Attribute Http.VB_VarHelpID = -1
Public RequestRef As Variant

''
' Delegate instance use to pass Events back.
'
' @property Delegate
' @type WebClientAsyncInstance
''
Public Delegate As WebClientAsyncInstance

' --------------------------------------------- '
' Windows Kernel timing functions
' --------------------------------------------- '

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function getFrequency Lib "kernel32" _
        Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function getFrequency Lib "kernel32" _
        Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
#End If

Public Sub ExecuteAsyncInstance()
    WebHelpers.LogDebug "Called", "WebClientAsyncInstance.ExecuteAsyncInstance"
    Set Me.Http = Client.PrepareHttpRequest(Me.Request)
    
    Me.Http.Send Request.Body
    RequestStartTicks = GetTickCount()
    WebHelpers.LogDebug "Ended", "WebAsyncRequestInstance.ExecuteAsyncInstance"
End Sub
Public Sub Abort(Optional Reason As String = "")
    WebHelpers.LogDebug "Called", "WebAsyncRequestInstance.Abort"
    If Reason <> "" Then Reason = ": " & Reason
    If Not Me.Http Is Nothing Then
        Me.Http.Abort
        WebHelpers.LogDebug "Aborted" & Reason, "WebClientAsyncInstance.Abort"
    Else
        WebHelpers.LogDebug "Not running" & Reason, "WebClientAsyncInstance.Abort"
    End If
    
    Set Me.Request = Nothing
    Set Me.Http = Nothing
    WebHelpers.LogDebug "Ended", "WebAsyncRequestInstance.Abort"
End Sub
Public Sub Http_OnError(ByVal ErrNo As Long, ByVal ErrDesc As String)
    WebHelpers.LogError "WinHttpRequest error" & ErrNo & ": " & ErrDesc, "WebClientAsyncInstance.Http_OnError"
    Delegate.Delegate_AsyncError Me.nonce, ErrNo, ErrDesc, Me.RequestRef
End Sub
Public Sub Delegate_AsyncError(ByVal nonce As String, ByVal ErrNo As Long, ByVal ErrDesc As String, RequestRef As Variant)
    RaiseEvent AsyncError(nonce, ErrNo, ErrDesc, RequestRef)
End Sub
Public Sub Http_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    WebHelpers.LogDebug "WinHttpRequest response starting: " & ElapsedTimeMs() & "ms", "WebClientAsyncInstance.Http_OnResponseStart"
End Sub
Private Sub Http_OnResponseDataAvailable(Data() As Byte)
    WebHelpers.LogDebug "WinHttpRequest response data available: " & ElapsedTimeMs() & "ms", "WebClientAsyncInstance.Http_OnResponseDataAvailable"
End Sub
Private Sub Http_OnResponseFinished()
    WebHelpers.LogDebug "WinHttpRequest response finished: " & ElapsedTimeMs() & "ms", "WebClientAsyncInstance.Http_OnResponseFinish"
    
    Dim web_Response As New WebResponse

    web_Response.CreateFromHttp Me.Client, Me.Request, Me.Http
    WebHelpers.LogResponse Me.Client, Me.Request, web_Response
    
    Set Me.Request = Nothing
    Set Me.Http = Nothing
    WebHelpers.LogDebug "Raising AsyncResponse Event: " & Me.nonce, "WebClientAsyncInstance.OnResponseFinish"
    Delegate.Delegate_AsyncResponse Me.nonce, web_Response, Me.RequestRef
End Sub
Public Sub Delegate_AsyncResponse(ByVal nonce As String, web_Response As WebResponse, RequestRef As Variant)
    RaiseEvent AsyncResponse(nonce, web_Response, RequestRef)
End Sub
Private Function ElapsedTimeMs() As Long
    ' Test whether you are using the 64-bit version of Office 2010.
    Dim e As Long
    
    ElapsedTimeMs = 0
    e = GetTickCount()
    If RequestStartTicks = 0 Then
        RequestStartTicks = e
    Else
        e = e - RequestStartTicks
        If e < 0 Then e = CLng(CCur(e) + 4294967296#)
        If TimerTickFrequency = 0 Then getFrequency TimerTickFrequency
        If TimerTickFrequency <> 0 Then ElapsedTimeMs = e / TimerTickFrequency
    End If
End Function
Private Sub Class_Terminate()
    WebHelpers.LogDebug "Terminating", "WebClientAsyncInstance.terminate"
    Set Me.Request = Nothing
    Set Me.Client = Nothing
    Set Me.Http = Nothing
End Sub

#End If

