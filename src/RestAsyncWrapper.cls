VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestAsyncWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' RestAsyncWrapper v4.0.0-beta.2
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Wrapper RestClient and RestRequest that enables callback-style async requests
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private pClient As RestClient

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Request As RestRequest
Public Callback As String
Public CallbackArgs As Variant
Public WithEvents Http As WinHttpRequest
Attribute Http.VB_VarHelpID = -1

Public Property Get Client() As RestClient
    Set Client = pClient
End Property
Public Property Set Client(Value As RestClient)
    If pClient Is Nothing Or Value Is Nothing Then
        Set pClient = Value
    Else
        ' If a Client is changed while other Requests are executing, it may introduce unexpected behavior
        ' Guard against changing Client and instead recommend creating a new AsyncWrapper per Client
        Err.Raise vbObjectError + 1, "RestAsyncWrapper", "The Client for a RestAsyncWrapper should not be changed as it may affect any currently executing Requests. " & _
            "A new RestAsyncWrapper should be created for each RestClient."
    End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Execute the specified request asynchronously
'
' @param {RestRequest} Request The request to execute
' @param {String} Callback Name of function to call when request completes
' @param {Variant} [CallbackArgs] Variable array of arguments that get passed directly to callback function
' @return {Boolean} Status of initiating request
' --------------------------------------------- '

Public Sub ExecuteAsync(Request As RestRequest, Callback As String, Optional ByVal CallbackArgs As Variant)
    ' In order for AsyncWrapper to be reusable, clone then execute with clone
    ' - AsyncWrapper can only watch one WinHttpRequest's events
    ' - Callback + CallbackArgs would need to be stored per Request
    Dim Async As RestAsyncWrapper
    Set Async = Me.Clone
    Async.PrepareAndExecuteRequest Request, Callback, CallbackArgs
    Exit Sub
End Sub

''
' Clone wrapper
' --------------------------------------------- '
Public Function Clone() As RestAsyncWrapper
    Set Clone = New RestAsyncWrapper
    Set Clone.Client = Me.Client
End Function

''
' Once everything has been prepared, execute request
'
' @param {RestRequest} Request
' @param {String} Callback
' @param {Variant} [CallbackArgs]
' --------------------------------------------- '
Public Sub PrepareAndExecuteRequest(Request As RestRequest, Callback As String, Optional ByVal CallbackArgs As Variant)
    Me.Callback = Callback
    Me.CallbackArgs = CallbackArgs
    
    Set Me.Request = Request.Clone
    Set Me.Http = Me.Client.PrepareHttpRequest(Request)

    StartTimeoutTimer
    Me.Http.Send Request.Body
    Exit Sub
    
ErrorHandling:
    
    Set Me.Http = Nothing
    Set Me.Request = Nothing
    
    ' Rethrow error
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''
' Handle timeouts
' --------------------------------------------- '
Public Sub TimedOut()
    ' Callback
    RestHelpers.LogDebug "Timed out", "RestAsyncWrapper.TimedOut"
    RunCallback RestHelpers.CreateResponse(WebStatusCode.RequestTimeout, "Request Timeout")
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub RunCallback(Response As RestResponse)
    ' Run callback function (needs to be a public function),
    ' passing in response and any defined callback arguments
    '
    ' callback({RestResponse})
    ' OR callback({RestResponse}, {Variant})
    '
    ' Example:
    ' Public Function Callback(response As RestResponse, args As Variant)
    '     Debug.Print "Callback: " & response.StatusCode
    '     For i = LBound(args) To UBound(args)
    '         Debug.Print args(i) & " was passed into async execute"
    '     Next i
    ' End Function
    
    RestHelpers.LogResponse Response, Me.Request
    
    If Not Me.Client.Authenticator Is Nothing Then
        Me.Client.Authenticator.AfterExecute Me.Client, Me.Request, Response
    End If
    If Me.Callback <> "" Then
        RestHelpers.LogDebug Me.Callback, "RestAsyncWrapper.RunCallback"
        If Not IsMissing(Me.CallbackArgs) Then
            Application.Run Me.Callback, Response, Me.CallbackArgs
        Else
            Application.Run Me.Callback, Response
        End If
    End If
    
    Set Me.Http = Nothing
    Set Me.Request = Nothing
    
End Sub

''
' Start timeout timer
' --------------------------------------------- '
Private Sub StartTimeoutTimer()
    RestHelpers.StartTimeoutTimer Me, Me.Client.TimeoutMS
End Sub

''
' Stop timeout timer
' --------------------------------------------- '
Private Sub StopTimeoutTimer()
    RestHelpers.StopTimeoutTimer Me
End Sub

''
' Process asynchronous requests
' --------------------------------------------- '
Private Sub Http_OnResponseFinished()
    StopTimeoutTimer
    
    ' Callback
    Dim Response As RestResponse
    Set Response = RestHelpers.CreateResponseFromHttp(Me.Http, Me.Request.ResponseFormat)
    RunCallback Response
End Sub

Private Sub Class_Terminate()
    Set Me.Client = Nothing
    Set Me.Request = Nothing
End Sub

