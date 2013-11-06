VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' RestRequest v1.1.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Create a request for use with a rest client
'
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const cacheBreakerKey As String = "cachebreaker"
Private pHeaders As Object
Private pParameters As Object
Private pQuerystringParams As Object
Private pUrlSegments As Object
Private pBody As Object
Private pBodyString As String
Private pContentType As String
Private pHttpRequest As Object

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum AvailableMethods
    httpGET
    httpPOST
    httpPUT
    httpDELETE
    httpPATCH
End Enum
Public Enum AvailableFormats
    json
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Resource As String
Public Method As AvailableMethods
Public Format As AvailableFormats
Public RootElement As String
Public Callback As String
Public BaseUrl As String
Public RequireHTTPS As Boolean
Public CallbackArgs As Variant
Public IncludeCacheBreaker As Boolean

Public Property Get Headers() As Object
    If pHeaders Is Nothing Then: Set pHeaders = CreateObject("Scripting.Dictionary")
    Set Headers = pHeaders
End Property
Public Property Let Headers(Value As Object)
    Set pHeaders = Value
End Property

Public Property Get Parameters() As Object
    If pParameters Is Nothing Then: Set pParameters = CreateObject("Scripting.Dictionary")
    Set Parameters = pParameters
End Property
Public Property Let Parameters(Value As Object)
    Set pParameters = Value
End Property

Public Property Get QuerystringParams() As Object
    If pQuerystringParams Is Nothing Then: Set pQuerystringParams = CreateObject("Scripting.Dictionary")
    Set QuerystringParams = pQuerystringParams
End Property
Public Property Let QuerystringParams(Value As Object)
    Set pQuerystringParams = Value
End Property

Public Property Get UrlSegments() As Object
    If pUrlSegments Is Nothing Then: Set pUrlSegments = CreateObject("Scripting.Dictionary")
    Set UrlSegments = pUrlSegments
End Property
Public Property Let UrlSegments(Value As Object)
    Set pUrlSegments = Value
End Property

Public Property Get FormattedResource() As String
    Dim segment As Variant
    Dim parameterKey As Variant
    
    FormattedResource = Me.Resource
    
    ' Replace url segments
    For Each segment In Me.UrlSegments.keys
        FormattedResource = Replace(FormattedResource, "{" & segment & "}", Me.UrlSegments(segment))
    Next segment
    FormattedResource = Replace(FormattedResource, "{format}", Me.FormatName())
    
    ' Add cache breaker to GET requests
    ' [MSXML2.XMLHTTP refresh](http://www.tek-tips.com/viewthread.cfm?qid=1563400)
    ' Looks like a last-modified could work here too, but may also require control over server response
    ' Also, with update to ServerXMLHTTP, neither may be needed
    If Me.Method = httpGET And Me.IncludeCacheBreaker Then
        Me.AddQuerystringParam cacheBreakerKey, timer() * 100
    End If
    
    ' Add querystring
    If (Me.Method = httpGET And Not Me.Parameters Is Nothing) Or Not Me.QuerystringParams Is Nothing Then
        If (Me.Parameters.count > 0 And Me.Method = httpGET) Or Me.QuerystringParams.count > 0 Then
            If InStr(FormattedResource, "?") <= 0 Then
                FormattedResource = FormattedResource & "?"
            End If
            
            ' Only load parameters to querystring if GET request (otherwise they are added to the body)
            If Me.Method = httpGET Then
                For Each parameterKey In Me.Parameters.keys()
                    If Right(FormattedResource, 1) <> "?" Then: FormattedResource = FormattedResource & "&"
                    FormattedResource = FormattedResource & URLEncode(parameterKey) & "=" & URLEncode(Me.Parameters(parameterKey))
                Next parameterKey
            End If
            
            For Each parameterKey In Me.QuerystringParams.keys()
                If Right(FormattedResource, 1) <> "?" Then: FormattedResource = FormattedResource & "&"
                FormattedResource = FormattedResource & URLEncode(parameterKey) & "=" & URLEncode(Me.QuerystringParams(parameterKey))
            Next parameterKey
        End If
    End If
End Property

Public Property Get Body() As String
    ' Add body if it's defined or parameters have been set and it is not a GET request
    If Not pBody Is Nothing Or pBodyString <> "" Or (Me.Parameters.count > 0 And Me.Method <> httpGET) Then
        Select Case Me.Format
        ' (Currently only JSON is supported)
        Case Else
            If pBodyString <> "" Then
                If Me.Parameters.count > 0 And Me.Method <> httpGET Then
                    Err.Raise vbObjectError + 1, "RestRequest.Body", "Unable to combine body string and parameters"
                Else
                    Body = pBodyString
                End If
            Else
                If Me.Method <> httpGET Then
                    ' Combine defined body and parameters and convert to JSON
                    Body = RestHelpers.ConvertToJSON(CombineObjects(pBody, Me.Parameters))
                Else
                    Body = RestHelpers.ConvertToJSON(pBody)
                End If
            End If
        End Select
    End If
End Property

Public Property Get FullUrl(Optional ClientBaseUrl As String = "") As String
    If Me.BaseUrl = "" Then Me.BaseUrl = ClientBaseUrl
    
    ' If protocol is missing from base url, add it based on RequireHTTPS property
    If Left(Me.BaseUrl, 7) <> "http://" And Left(Me.BaseUrl, 8) <> "https://" Then
        If Me.RequireHTTPS Then
            Me.BaseUrl = "https://" & Me.BaseUrl
        Else
            Me.BaseUrl = "http://" & Me.BaseUrl
        End If
    End If
    
    FullUrl = JoinUrl(Me.BaseUrl, Me.FormattedResource)
End Property

Public Property Get MethodName() As String
    Select Case Me.Method
    Case httpDELETE
        MethodName = "DELETE"
    Case httpPUT
        MethodName = "PUT"
    Case httpPATCH
        MethodName = "PATCH"
    Case httpPOST
        MethodName = "POST"
    Case Else
        MethodName = "GET"
    End Select
End Property

Public Property Get FormatName() As String
    Select Case Me.Format
    Case Else
        FormatName = "json"
    End Select
End Property

Public Property Get ContentType() As String
    If pContentType <> "" Then
        ContentType = pContentType
    Else
        Select Case Me.Format
        Case Else
            If Me.Method <> httpGET And Me.Parameters.count > 0 Then
                ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
            Else
                ContentType = "application/json"
            End If
        End Select
    End If
End Property
Public Property Let ContentType(Value As String)
    pContentType = Value
End Property

Public Property Get HttpRequest() As Object
    Set HttpRequest = pHttpRequest
End Property
Public Property Set HttpRequest(ByVal Value As Object)
    Set pHttpRequest = Value
End Property


' ============================================= '
' Public Methods
' ============================================= '

''
' Add the specified key/value to the request header
'
' @param {String} key
' @param {String} value
' --------------------------------------------- '

Public Sub AddHeader(Key As String, Value As String)
    If Not Me.Headers.Exists(Key) Then
        Me.Headers.Add Key, Value
    Else
        Me.Headers(Key) = Value
    End If
End Sub

''
' Replace the defined segment in the url with the provided value
'
' @param {String} key
' @param {String} value
' --------------------------------------------- '

Public Sub AddUrlSegment(segment As String, Value As String)
    If Not Me.UrlSegments.Exists(segment) Then
        Me.UrlSegments.Add segment, Value
    Else
        Me.UrlSegments(segment) = Value
    End If
End Sub

''
' Add parameter to request (querystring for GET, body for POST)
'
' @param {String} key
' @param {Variant} value
' --------------------------------------------- '

Public Sub AddParameter(Key As String, Value As Variant)
    If Not Me.Parameters.Exists(Key) Then
        Me.Parameters.Add Key, Value
    Else
        Me.Parameters(Key) = Value
    End If
End Sub

''
' Add querystring parameter to request
'
' @param {String} key
' @param {Variant} value
' --------------------------------------------- '

Public Sub AddQuerystringParam(Key As String, Value As Variant)
    If Not Me.QuerystringParams.Exists(Key) Then
        Me.QuerystringParams.Add Key, Value
    Else
        Me.QuerystringParams(Key) = Value
    End If
End Sub

''
' Add body to request
'
' @param {Dictionary} bodyVal Object to add to body (will be converted to string)
' --------------------------------------------- '

Public Function AddBody(bodyVal As Object)
    Set pBody = bodyVal
End Function

''
' Add body as string
'
' @param {String} bodyVal
' --------------------------------------------- '

Public Function AddBodyString(bodyVal As String)
    pBodyString = bodyVal
End Function

''
' Helper for creating response
'
' @param {StatusCodes} StatusCode
' @param {String} StatusDescription
' @param {String} [Content=""]
' @return {RestResponse}
' --------------------------------------------- '

Public Function CreateResponse(StatusCode As StatusCodes, StatusDescription As String, Optional Content As String = "") As RestResponse
    Set CreateResponse = New RestResponse
    CreateResponse.StatusCode = StatusCode
    CreateResponse.StatusDescription = StatusDescription
    CreateResponse.Content = Content
    
    Select Case Me.Format
    Case Else
        Set CreateResponse.Data = RestHelpers.ParseJSON(Content)
    End Select
End Function

''
' Helper for creating response
'
' @param {XMLHttpRequest} Http
' @return {RestResponse}
' --------------------------------------------- '

Public Function CreateResponseFromHttp(Http As Object) As RestResponse
    Set CreateResponseFromHttp = Me.CreateResponse(Http.Status, Http.StatusText, Http.ResponseText)
End Function

''
' Process asynchronous requests
'
' --------------------------------------------- '

Public Sub ReadyStateChangeHandler()
Attribute ReadyStateChangeHandler.VB_UserMemId = 0
    ' During development, uncomment to watch changes in ready state
    ' Debug.Print "ReadyStageChangeHandler: " & Me.HttpRequest.readyState

    ' ReadyStates [The XMLHttpRequest Object](http://www.w3.org/TR/2006/WD-XMLHttpRequest-20060405/):
    ' 0. Uninintialized
    ' 1. Open
    ' 2. Sent
    ' 3. Receiving
    ' 4. Loaded
    
    ' Once loaded, process result
    If Me.HttpRequest.readyState = 4 Then
        Me.StopTimeoutTimer
        
        ' Callback
        RunCallback Me.CreateResponseFromHttp(Me.HttpRequest)
        
        ' Request is finished, clean up
        Set Me.HttpRequest = Nothing
    End If
    
    ' TODO: Check for timeout and cleanup
End Sub

''
' Start timeout timer
'
' --------------------------------------------- '

Public Sub StartTimeoutTimer(TimeoutMS As Integer)
    RestHelpers.StartTimeoutTimer Me, TimeoutMS
End Sub

''
' Stop timeout timer
'
' --------------------------------------------- '

Public Sub StopTimeoutTimer()
    RestHelpers.StopTimeoutTimer Me
End Sub

''
' Handle timeouts
'
' --------------------------------------------- '

Public Sub TimedOut()
    ' Callback
    RunCallback Me.CreateResponse(StatusCodes.GatewayTimeout, "Gateway Timeout")
    
    ' Request is finished, clean up
    Set Me.HttpRequest = Nothing
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
    '
    If Me.Callback <> "" Then
        If Not IsMissing(Me.CallbackArgs) Then
            Application.Run Me.Callback, Response, Me.CallbackArgs
        Else
            Application.Run Me.Callback, Response
        End If
    End If
End Sub

Private Sub Class_Initialize()
    ' Set default values
    Me.IncludeCacheBreaker = True
End Sub

Private Sub Class_Terminate()
    ' Clean up
    If Not Me.HttpRequest Is Nothing Then: Set Me.HttpRequest = Nothing
End Sub
