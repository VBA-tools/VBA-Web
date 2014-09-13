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
' RestRequest v3.1.4
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Create a request for use with a rest client
'
' @dependencies: Microsoft Scripting Runtime
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private pHeaders As Dictionary
Private pParameters As Dictionary
Private pQuerystringParams As Dictionary
Private pUrlSegments As Dictionary
Private pCookies As Dictionary
Private pBody As Variant
Private pBodyString As String
Private pContentType As String
Private pAccept As String
Private pContentLength As Long
Private pHttpRequest As Object
Private pId As String

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
    plaintext
    json
    formurlencoded
    xml
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Resource As String
Public Method As AvailableMethods
Public RequestFormat As AvailableFormats
Public ResponseFormat As AvailableFormats
Public Callback As String
Public BaseUrl As String
Public RequireHTTPS As Boolean
Public CallbackArgs As Variant
Public IncludeContentLength As Boolean
Public Authenticator As IAuthenticator
Public Client As RestClient

Public Property Get Format() As AvailableFormats
    Format = RequestFormat
End Property
Public Property Let Format(Value As AvailableFormats)
    Me.RequestFormat = Value
    Me.ResponseFormat = Value
End Property

Public Property Get Headers() As Dictionary
    If pHeaders Is Nothing Then: Set pHeaders = New Dictionary
    Set Headers = pHeaders
End Property
Public Property Set Headers(Value As Dictionary)
    Set pHeaders = Value
End Property

Public Property Get Parameters() As Dictionary
    If pParameters Is Nothing Then: Set pParameters = New Dictionary
    Set Parameters = pParameters
End Property
Public Property Set Parameters(Value As Dictionary)
    Set pParameters = Value
End Property

Public Property Get QuerystringParams() As Dictionary
    If pQuerystringParams Is Nothing Then: Set pQuerystringParams = New Dictionary
    Set QuerystringParams = pQuerystringParams
End Property
Public Property Set QuerystringParams(Value As Dictionary)
    Set pQuerystringParams = Value
End Property

Public Property Get UrlSegments() As Dictionary
    If pUrlSegments Is Nothing Then: Set pUrlSegments = New Dictionary
    Set UrlSegments = pUrlSegments
End Property
Public Property Set UrlSegments(Value As Dictionary)
    Set pUrlSegments = Value
End Property

Public Property Get Cookies() As Dictionary
    If pCookies Is Nothing Then: Set pCookies = New Dictionary
    Set Cookies = pCookies
End Property
Public Property Set Cookies(Value As Dictionary)
    Set pCookies = Value
End Property

Public Property Let IncludeCacheBreaker(Value As Boolean)
    Debug.Print "Excel-REST: DEPRECATED IncludeCacheBreaker has been deprecated since it is not required with ServerXMLHTTP. It will be removed in Excel-REST v4"
End Property

Public Property Get FormattedResource() As String
    Dim Segment As Variant
    
    FormattedResource = Me.Resource
    
    ' Replace url segments
    For Each Segment In Me.UrlSegments.Keys
        FormattedResource = Replace(FormattedResource, "{" & Segment & "}", Me.UrlSegments(Segment))
    Next Segment
    FormattedResource = Replace(FormattedResource, "{format}", RestHelpers.FormatToName(Me.ResponseFormat))
    
    ' Add querystring
    If (Me.Method = httpGET And Not Me.Parameters Is Nothing) Or Not Me.QuerystringParams Is Nothing Then
        If (Me.Parameters.count > 0 And Me.Method = httpGET) Or Me.QuerystringParams.count > 0 Then
            If InStr(FormattedResource, "?") <= 0 Then
                FormattedResource = FormattedResource & "?"
            Else
                FormattedResource = FormattedResource & "&"
            End If
            
            ' Only load parameters to querystring if GET request (otherwise they are added to the body)
            If Me.Method = httpGET Then
                FormattedResource = FormattedResource & RestHelpers.ConvertToUrlEncoded(RestHelpers.CombineObjects(Me.Parameters, Me.QuerystringParams))
            Else
                FormattedResource = FormattedResource & RestHelpers.ConvertToUrlEncoded(Me.QuerystringParams)
            End If
        End If
    End If
End Property

Public Property Get Body() As String
    ' Add body if it's defined or parameters have been set and it is not a GET request
    If Not IsEmpty(pBody) Or pBodyString <> "" Or (Me.Parameters.count > 0 And Me.Method <> httpGET) Then
        If pBodyString <> "" Then
            If Me.Parameters.count > 0 And Me.Method <> httpGET Then
                Err.Raise vbObjectError + 1, "RestRequest.Body", "Unable to combine body string and parameters"
            Else
                Body = pBodyString
            End If
        Else
            Dim BodyValue As Variant
            If RestHelpers.IsArray(pBody) And Me.Parameters.count > 0 And Me.Method <> httpGET Then
                Err.Raise vbObjectError + 1, "RestRequest.Body", "Unable to combine body array and parameters"
            ElseIf Not RestHelpers.IsArray(pBody) And Me.Parameters.count > 0 And Me.Method <> httpGET Then
                If Me.Parameters.count > 0 And Not IsEmpty(pBody) Then
                    Set BodyValue = CombineObjects(Me.Parameters, pBody)
                ElseIf Me.Parameters.count > 0 Then
                    Set BodyValue = Me.Parameters
                Else
                    Set BodyValue = pBody
                End If
            ElseIf VarType(pBody) = vbObject Then
                Set BodyValue = pBody
            Else
                BodyValue = pBody
            End If
            
            If Me.RequestFormat <> AvailableFormats.plaintext Then
                Body = RestHelpers.ConvertToFormat(BodyValue, Me.RequestFormat)
            Else
                Body = BodyValue
            End If
        End If
    End If
End Property

Public Property Get FullUrl(Optional ClientBaseUrl As String = "") As String
    Dim Formatted As String
    Formatted = Me.FormattedResource
    If Me.BaseUrl = "" Then Me.BaseUrl = ClientBaseUrl
    
    ' If protocol is missing, add it based on RequireHTTPS property
    If RestHelpers.IncludesProtocol(Me.BaseUrl) = "" And RestHelpers.IncludesProtocol(Formatted) = "" Then
        If Me.RequireHTTPS Then
            Me.BaseUrl = "https://" & Me.BaseUrl
        Else
            Me.BaseUrl = "http://" & Me.BaseUrl
        End If
    End If
    
    ' If protocol is included on base and resource, use base
    ' If protocol is included on resource and base is not blank, move to base
    If RestHelpers.IncludesProtocol(Me.BaseUrl) <> "" And RestHelpers.IncludesProtocol(Formatted) <> "" Then
        Formatted = RestHelpers.RemoveProtocol(Formatted)
    ElseIf RestHelpers.IncludesProtocol(Formatted) <> "" And Me.BaseUrl <> "" Then
        Me.BaseUrl = RestHelpers.IncludesProtocol(Formatted) & Me.BaseUrl
        Formatted = RestHelpers.RemoveProtocol(Formatted)
    End If
    
    FullUrl = JoinUrl(Me.BaseUrl, Formatted)
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
    FormatName = RestHelpers.FormatToName(Me.RequestFormat)
End Property

Public Property Get ContentType() As String
    If pContentType <> "" Then
        ContentType = pContentType
    Else
        ContentType = RestHelpers.FormatToContentType(Me.RequestFormat)
    End If
End Property
Public Property Let ContentType(Value As String)
    pContentType = Value
End Property

Public Property Get Accept() As String
    If pAccept <> "" Then
        Accept = pAccept
    Else
        Accept = RestHelpers.FormatToContentType(Me.ResponseFormat)
    End If
End Property
Public Property Let Accept(Value As String)
    pAccept = Value
End Property

Public Property Get ContentLength() As Long
    If pContentLength > 0 Then
        ContentLength = pContentLength
    Else
        ContentLength = Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(Value As Long)
    pContentLength = Value
End Property

Public Property Get HttpRequest() As Object
    Set HttpRequest = pHttpRequest
End Property
Public Property Set HttpRequest(ByVal Value As Object)
    Set pHttpRequest = Value
End Property

Public Property Get Id() As String
    If pId = "" Then
        On Error Resume Next
        ' First, try creating GUID
        pId = Replace(Replace(Left(CreateObject("Scriptlet.TypeLibz").GUID, 38), "{", ""), "}", "")
        
        If pId = "" Then
            ' If that fails, use nonce
            pId = UCase(RestHelpers.CreateNonce)
        End If
    End If
    
    Id = pId
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
    RestHelpers.AddToDictionary Me.Headers, Key, Value
End Sub

''
' Replace the defined segment in the url with the provided value
'
' @param {String} key
' @param {String} value
' --------------------------------------------- '
Public Sub AddUrlSegment(Segment As String, Value As String)
    RestHelpers.AddToDictionary Me.UrlSegments, Segment, Value
End Sub

''
' Add parameter to request (querystring for GET, body for POST)
'
' @param {String} key
' @param {Variant} value
' --------------------------------------------- '
Public Sub AddParameter(Key As String, Value As Variant)
    RestHelpers.AddToDictionary Me.Parameters, Key, Value
End Sub

''
' Add querystring parameter to request
'
' @param {String} key
' @param {Variant} value
' --------------------------------------------- '
Public Sub AddQuerystringParam(Key As String, Value As Variant)
    RestHelpers.AddToDictionary Me.QuerystringParams, Key, Value
End Sub

''
' Add cookie to request
'
' @param {String} key
' @param {Variant} value
' --------------------------------------------- '
Public Sub AddCookie(Key As String, Value As Variant)
    RestHelpers.AddToDictionary Me.Cookies, Key, Value
End Sub

''
' Add body to request
'
' @param {Variant} bodyVal Object/Collection/Array to add to body (will be converted to string)
' --------------------------------------------- '
Public Sub AddBody(BodyVal As Variant)
    If IsObject(BodyVal) Then
        Set pBody = BodyVal
    ElseIf RestHelpers.IsArray(BodyVal) Then
        pBody = BodyVal
    Else
        pBodyString = BodyVal
    End If
End Sub

''
' Add body as string
'
' @param {String} bodyVal
' --------------------------------------------- '
Public Sub AddBodyString(BodyVal As String)
    pBodyString = BodyVal
End Sub

''
' Process asynchronous requests
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
        Dim Response As RestResponse
        Set Response = RestHelpers.CreateResponseFromHttp(Me.HttpRequest, Me.ResponseFormat)
        RestHelpers.LogResponse Response, Me
        RunCallback Response
    End If
End Sub

''
' Start timeout timer
' --------------------------------------------- '
Public Sub StartTimeoutTimer(TimeoutMS As Long)
    RestHelpers.StartTimeoutTimer Me, TimeoutMS
End Sub

''
' Stop timeout timer
' --------------------------------------------- '
Public Sub StopTimeoutTimer()
    RestHelpers.StopTimeoutTimer Me
End Sub

''
' Handle timeouts
' --------------------------------------------- '
Public Sub TimedOut()
    ' Callback
    RestHelpers.LogDebug "Timed out", "RestRequest.TimedOut"
    RunCallback RestHelpers.CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
End Sub

''
' Clone request
'
' @return {RestRequest}
' --------------------------------------------- '
Public Function Clone() As RestRequest
    Set Clone = New RestRequest
    
    Clone.Resource = Me.Resource
    Clone.Method = Me.Method
    Clone.Accept = Me.Accept
    Clone.ContentType = Me.ContentType
    Clone.RequestFormat = Me.RequestFormat
    Clone.ResponseFormat = Me.ResponseFormat
    Clone.Callback = Me.Callback
    Clone.ContentLength = Me.ContentLength
    Clone.BaseUrl = Me.BaseUrl
    Clone.RequireHTTPS = Me.RequireHTTPS
    Clone.IncludeContentLength = Me.IncludeContentLength
    
    Set Clone.Headers = RestHelpers.CloneDictionary(Me.Headers)
    Set Clone.Parameters = RestHelpers.CloneDictionary(Me.Parameters)
    Set Clone.QuerystringParams = RestHelpers.CloneDictionary(Me.QuerystringParams)
    Set Clone.UrlSegments = RestHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = RestHelpers.CloneDictionary(Me.Cookies)
    
    If pBodyString <> "" Then
        Clone.AddBody pBodyString
    Else
        Clone.AddBody pBody
    End If
End Function

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
    
    If Not Me.Authenticator Is Nothing Then
        ' Don't pass Authenticator and Client with Request to AfterExecute
        Dim Auth As IAuthenticator
        Dim Client As RestClient
        Set Auth = Me.Authenticator
        Set Client = Me.Client
        Set Me.Authenticator = Nothing
        Set Me.Client = Nothing
        
        Auth.AfterExecute Client, Me, Response
    End If
    If Me.Callback <> "" Then
        RestHelpers.LogDebug Me.Callback, "RestRequest.RunCallback"
        If Not IsMissing(Me.CallbackArgs) Then
            Application.Run Me.Callback, Response, Me.CallbackArgs
        Else
            Application.Run Me.Callback, Response
        End If
    End If

    ' Request is finished, clean up
    Set Me.HttpRequest = Nothing
End Sub

Private Sub Class_Initialize()
    ' Set default values
    Me.IncludeContentLength = True
    Me.RequestFormat = json
    Me.ResponseFormat = json
End Sub

Private Sub Class_Terminate()
    ' Clean up
    If Not Me.HttpRequest Is Nothing Then: Set Me.HttpRequest = Nothing
    If Not Me.Authenticator Is Nothing Then: Set Me.Authenticator = Nothing
    If Not Me.Client Is Nothing Then: Set Me.Client = Nothing
End Sub
