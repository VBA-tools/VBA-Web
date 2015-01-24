VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebAsyncWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebAsyncWrapper v4.0.0-rc.5
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Wrapper WebClient and WebRequest that enables callback-style async requests
'
' _Note_ Windows-only and requires reference to "Microsoft WinHTTP Services, version 5.1"
'
' Errors:
' 11050 / 80042b2a / -2147210454 - Client should not be changed
'
' @class WebAsyncWrapper
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private web_pClient As WebClient

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Request that is currently executing.
'
' @property Request
' @type WebRequest
''
Public Request As WebRequest

''
' Function to call with response when request has completed.
'
' @property Callback
' @type String
''
Public Callback As String

''
' Array of arguments to pass to callback along with response
'
' @property CallbackArgs
' @type Variant
''
Public CallbackArgs As Variant

''
' @property Http
' @type WebHttpRequest
''
Public WithEvents Http As WinHttpRequest
Attribute Http.VB_VarHelpID = -1

''
' Client used for executing requests
'
' @property Client
' @type WebClient
' @throws 11050 / 80042b2a / -2147210454 - Client should not be changed
''
Public Property Get Client() As WebClient
    Set Client = web_pClient
End Property
Public Property Set Client(Value As WebClient)
    If web_pClient Is Nothing Or Value Is Nothing Then
        Set web_pClient = Value
    Else
        ' If a Client is changed while other Requests are executing, it may introduce unexpected behavior
        ' Guard against changing Client and instead recommend creating a new AsyncWrapper per Client
        Dim web_ErrorDescription As String
        web_ErrorDescription = "The Client for a WebAsyncWrapper should not be changed as it may affect any currently executing Requests. " & _
            "A new WebAsyncWrapper should be created for each WebClient."
        
        WebHelpers.LogError web_ErrorDescription, "WebAsyncWrapper.Client", vbObjectError + 11050
        Err.Raise vbObjectError + 11050, "WebAsyncWrapper.Client", web_ErrorDescription
    End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Execute the specified request asynchronously
'
' @method ExecuteAsync
' @param {WebRequest} Request The request to execute
' @param {String} Callback Name of function to call when request completes
' @param {Variant} [CallbackArgs] Variable array of arguments that get passed directly to callback function
''
Public Sub ExecuteAsync(Request As WebRequest, Callback As String, Optional ByVal CallbackArgs As Variant)
    ' In order for AsyncWrapper to be reusable, clone then execute with clone
    ' - AsyncWrapper can only watch one WinHttpRequest's events
    ' - Callback + CallbackArgs would need to be stored per Request
    Dim web_Async As WebAsyncWrapper
    
    Set web_Async = Me.Clone
    web_Async.PrepareAndExecuteRequest Request, Callback, CallbackArgs
End Sub

''
' Clone wrapper
'
' @internal
' @method Clone
' @return WebAsyncWrapper
''
Public Function Clone() As WebAsyncWrapper
    Set Clone = New WebAsyncWrapper
    Set Clone.Client = Me.Client
End Function

''
' Once everything has been prepared, execute request
'
' @internal
' @method PrepareAndExecuteRequest
' @param {RestRequest} Request
' @param {String} Callback
' @param {Variant} [CallbackArgs]
''
Public Sub PrepareAndExecuteRequest(Request As WebRequest, Callback As String, Optional ByVal CallbackArgs As Variant)
    On Error GoTo web_ErrorHandling

    Me.Callback = Callback
    Me.CallbackArgs = CallbackArgs
    
    Set Me.Request = Request.Clone
    Set Me.Http = Me.Client.PrepareHttpRequest(Request)

    web_StartTimeoutTimer
    Me.Http.Send Request.Body
    Exit Sub
    
web_ErrorHandling:
    
    Set Me.Http = Nothing
    Set Me.Request = Nothing
    
    ' Rethrow error
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''
' Handle timeout
'
' @internal
' @method TimedOut
''
Public Sub TimedOut()
    Dim web_Response As New WebResponse
    web_Response.StatusCode = WebStatusCode.RequestTimeout
    web_Response.StatusDescription = "Request Timeout"

    ' Callback
    WebHelpers.LogDebug "Timed out", "WebAsyncWrapper.TimedOut"
    web_RunCallback web_Response
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub web_RunCallback(web_Response As WebResponse)
    ' Run callback function (needs to be a public function),
    ' passing in response and any defined callback arguments
    '
    ' callback({RestResponse})
    ' OR callback({RestResponse}, {Variant})
    '
    ' Example:
    ' Public Function Callback(Response As WebResponse, Args As Variant)
    '     Debug.Print "Callback: " & response.StatusCode
    '     For i = LBound(args) To UBound(args)
    '         Debug.Print args(i) & " was passed into async execute"
    '     Next i
    ' End Function
    
    WebHelpers.LogResponse Me.Client, Me.Request, web_Response
    
    If Not Me.Client.Authenticator Is Nothing Then
        Me.Client.Authenticator.AfterExecute Me.Client, Me.Request, web_Response
    End If
    If Me.Callback <> "" Then
        WebHelpers.LogDebug Me.Callback, "WebAsyncWrapper.RunCallback"
        If Not IsMissing(Me.CallbackArgs) Then
            Application.Run Me.Callback, web_Response, Me.CallbackArgs
        Else
            Application.Run Me.Callback, web_Response
        End If
    End If
    
    Set Me.Http = Nothing
    Set Me.Request = Nothing
End Sub

' Start timeout timer
Private Sub web_StartTimeoutTimer()
    WebHelpers.StartTimeoutTimer Me, Me.Client.TimeoutMs
End Sub

' Stop timeout timer
Private Sub web_StopTimeoutTimer()
    WebHelpers.StopTimeoutTimer Me
End Sub

' Process asynchronous requests
Private Sub Http_OnResponseFinished()
    Dim web_Response As New WebResponse
    
    web_StopTimeoutTimer
    
    ' Callback
    web_Response.CreateFromHttp Me.Client, Me.Request, Me.Http
    web_RunCallback web_Response
End Sub

Private Sub Class_Terminate()
    Set Me.Client = Nothing
    Set Me.Request = Nothing
End Sub

