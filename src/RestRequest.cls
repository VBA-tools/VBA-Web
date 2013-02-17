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
' RestRequest v0.9.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Create a request for use with a rest client
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const cacheBreakerKey As String = "cachebreaker"
Private pHeaders As Dictionary
Private pParameters As Dictionary
Private pQuerystringParams As Dictionary
Private pUrlSegments As Dictionary
Private pBody As Dictionary
Private pHttpRequest As MSXML2.ServerXMLHTTP

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

Public Property Get Headers() As Dictionary
    If pHeaders Is Nothing Then: Set pHeaders = New Dictionary
    Set Headers = pHeaders
End Property
Public Property Let Headers(value As Dictionary)
    Set pHeaders = value
End Property

Public Property Get Parameters() As Dictionary
    If pParameters Is Nothing Then: Set pParameters = New Dictionary
    Set Parameters = pParameters
End Property
Public Property Let Parameters(value As Dictionary)
    Set pParameters = value
End Property

Public Property Get QuerystringParams() As Dictionary
    If pQuerystringParams Is Nothing Then: Set pQuerystringParams = New Dictionary
    Set QuerystringParams = pQuerystringParams
End Property
Public Property Let QuerystringParams(value As Dictionary)
    Set pQuerystringParams = value
End Property

Public Property Get UrlSegments() As Dictionary
    If pUrlSegments Is Nothing Then: Set pUrlSegments = New Dictionary
    Set UrlSegments = pUrlSegments
End Property
Public Property Let UrlSegments(value As Dictionary)
    Set pUrlSegments = value
End Property

Public Property Get FormattedResource(Optional includeCacheBreaker As Boolean = True) As String
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
    If Me.Method = httpGET And includeCacheBreaker Then
        If Me.Parameters Is Nothing Then
            Me.Parameters = New Dictionary
        End If
        Me.Parameters(cacheBreakerKey) = Timer() * 100
    End If
    
    ' Add querystring
    If (Me.Method = httpGET And Not Me.Parameters Is Nothing) Or Not Me.QuerystringParams Is Nothing Then
        If Me.Parameters.count > 0 Or Me.QuerystringParams.count > 0 Then
            FormattedResource = FormattedResource & "?"
            
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
    If Not pBody Is Nothing Or (Not Me.Parameters Is Nothing And Me.Method <> httpGET) Then
        Select Case Me.Format
        ' (Currently only JSON is supported)
        Case Else
            ' Combine defined body and parameters and convert to JSON
            Body = RestHelpers.ConvertToJSON(CombineObjects(pBody, Me.Parameters))
        End Select
    End If
End Property

Public Property Get FullUrl(Optional ClientBaseUrl As String = "") As String
    If Me.BaseUrl = "" Then Me.BaseUrl = ClientBaseUrl
    If Right(Me.BaseUrl, 1) <> "/" Then
        Me.BaseUrl = Me.BaseUrl & "/"
    End If
    
    ' If protocol is missing from base url, add it based on RequireHTTPS property
    If Left(Me.BaseUrl, 7) <> "http://" And Left(Me.BaseUrl, 8) <> "https://" Then
        If Me.RequireHTTPS Then
            Me.BaseUrl = "https://" & Me.BaseUrl
        Else
            Me.BaseUrl = "http://" & Me.BaseUrl
        End If
    End If
    
    FullUrl = Me.BaseUrl & Me.FormattedResource
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
    Select Case Me.Format
    Case Else
        If Me.Method <> httpGET And Me.Parameters.count > 0 Then
            ContentType = "application/x-www-form-urlencoded"
        Else
            ContentType = "application/json"
        End If
    End Select
End Property

Public Property Get HttpRequest() As MSXML2.ServerXMLHTTP
    Set HttpRequest = pHttpRequest
End Property
Public Property Let HttpRequest(ByVal value As MSXML2.ServerXMLHTTP)
    Set pHttpRequest = value
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

Public Sub AddHeader(key As String, value As String)
    If Not Me.Headers.Exists(key) Then
        Me.Headers.Add key, value
    Else
        Me.Headers(key) = value
    End If
End Sub

''
' Replace the defined segment in the url with the provided value
'
' @param {String} key
' @param {String} value
' --------------------------------------------- '

Public Sub AddUrlSegment(segment As String, value As String)
    If Not Me.UrlSegments.Exists(segment) Then
        Me.UrlSegments.Add segment, value
    Else
        Me.UrlSegments(segment) = value
    End If
End Sub

''
' Add parameter to request (querystring for GET, body for POST)
'
' @param {String} key
' @param {String} value
' --------------------------------------------- '

Public Sub AddParameter(key As String, value As String)
    If Not Me.Parameters.Exists(key) Then
        Me.Parameters.Add key, value
    Else
        Me.Parameters(key) = value
    End If
End Sub

''
' Add querystring parameter to request
'
' @param {String} key
' @param {String} value
' --------------------------------------------- '

Public Sub AddQuerystringParam(key As String, value As String)
    If Not Me.QuerystringParams.Exists(key) Then
        Me.QuerystringParams.Add key, value
    Else
        Me.QuerystringParams(key) = value
    End If
End Sub

''
' Add body to request
'
' @param {Dictionary} bodyVal Object to add to body (will be converted to string)
' --------------------------------------------- '

Public Function AddBody(bodyVal As Dictionary)
    Set pBody = bodyVal
End Function

''
' Process asynchronous requests
'
' --------------------------------------------- '

Public Sub ReadyStageChangeHandler()
Attribute ReadyStageChangeHandler.VB_UserMemId = 0
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
        ' Handle response
        Dim response As New RestResponse
        response.StatusCode = Me.HttpRequest.Status
        response.StatusDescription = Me.HttpRequest.statusText
        response.Content = Me.HttpRequest.responseText
        
        ' Serialize the response content
        Select Case Me.Format
        ' Currently, JSON is the only supported format
        Case Else
            Set response.Data = RestHelpers.ParseJSON(response.Content)
        End Select
        
        ' Callback
        If Me.Callback <> "" Then
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
            Application.Run Me.Callback, response, Me.CallbackArgs
        End If
        
        ' Request is finished, clean up
        Set Me.HttpRequest = Nothing
    End If
    
    ' TODO: Check for timeout and cleanup
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Terminate()
    ' Clean up
    If Not Me.HttpRequest Is Nothing Then: Set Me.HttpRequest = Nothing
End Sub
