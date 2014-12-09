VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebRequest v4.0.0-beta.5
' (c) Tim Hall - https://github.com/timhall/VBA-Web
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

Private pRequestFormat As WebFormat
Private pResponseFormat As WebFormat
Private pCustomRequestFormat As String
Private pCustomResponseFormat As String
Private pBody As Variant
Private pConvertedBody As String
Private pContentType As String
Private pAccept As String
Private pContentLength As Long
Private pId As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Resource As String
Public Method As WebMethod
Public Headers As Collection
Public QuerystringParams As Collection
Public UrlSegments As Dictionary
Public Cookies As Collection

Public Property Get Format() As WebFormat
    Format = RequestFormat
End Property
Public Property Let Format(Value As WebFormat)
    Me.RequestFormat = Value
    Me.ResponseFormat = Value
End Property

Public Property Get RequestFormat() As WebFormat
    RequestFormat = pRequestFormat
End Property
Public Property Let RequestFormat(Value As WebFormat)
    pRequestFormat = Value
    pConvertedBody = ""
End Property

Public Property Get ResponseFormat() As WebFormat
    ResponseFormat = pResponseFormat
End Property
Public Property Let ResponseFormat(Value As WebFormat)
    pResponseFormat = Value
    pConvertedBody = ""
End Property

Public Property Get CustomRequestFormat() As String
    CustomRequestFormat = pCustomRequestFormat
End Property
Public Property Let CustomRequestFormat(Value As String)
    pCustomRequestFormat = Value
    pConvertedBody = ""
    
    If Value <> "" Then
        pRequestFormat = WebFormat.Custom
    End If
End Property

Public Property Get CustomResponseFormat() As String
    CustomResponseFormat = pCustomResponseFormat
End Property
Public Property Let CustomResponseFormat(Value As String)
    pCustomResponseFormat = Value
    pConvertedBody = ""
    
    If Value <> "" Then
        ResponseFormat = WebFormat.Custom
    End If
End Property

Public Property Get ContentType() As String
    If pContentType <> "" Then
        ContentType = pContentType
    Else
        ContentType = WebHelpers.FormatToMediaType(Me.RequestFormat, Me.CustomRequestFormat)
    End If
End Property
Public Property Let ContentType(Value As String)
    pContentType = Value
End Property

Public Property Get Accept() As String
    If pAccept <> "" Then
        Accept = pAccept
    Else
        Accept = WebHelpers.FormatToMediaType(Me.ResponseFormat, Me.CustomResponseFormat)
    End If
End Property
Public Property Let Accept(Value As String)
    pAccept = Value
End Property

Public Property Get ContentLength() As Long
    If pContentLength >= 0 Then
        ContentLength = pContentLength
    Else
        ContentLength = Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(Value As Long)
    pContentLength = Value
End Property

Public Property Get Body() As Variant
    If Not IsEmpty(pBody) Then
        If VarType(pBody) = vbString Then
            Body = pBody
        ElseIf pConvertedBody = "" Then
            Body = WebHelpers.ConvertToFormat(pBody, Me.RequestFormat, Me.CustomRequestFormat)
            pConvertedBody = Body
        Else
            Body = pConvertedBody
        End If
    End If
End Property
Public Property Let Body(Value As Variant)
    ' Clear existing converted body
    pConvertedBody = ""
    
    pBody = Value
End Property
Public Property Set Body(Value As Variant)
    ' Clear existing converted body
    pConvertedBody = ""
    
    Set pBody = Value
End Property

Public Property Get FormattedResource() As String
    FormattedResource = Me.Resource
    
    ' Replace url segments
    Dim web_Segment As Variant
    For Each web_Segment In Me.UrlSegments.Keys
        FormattedResource = Replace(FormattedResource, "{" & web_Segment & "}", Me.UrlSegments(web_Segment))
    Next web_Segment
    
    ' Add querystring
    If Me.QuerystringParams.Count > 0 Then
        If VBA.InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If
        
        FormattedResource = FormattedResource & WebHelpers.ConvertToUrlEncoded(Me.QuerystringParams)
    End If
End Property

Public Property Get Id() As String
    If pId = "" Then: pId = WebHelpers.CreateNonce
    Id = pId
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add the specified key/value to the request headers
'
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddHeader(Key As String, Value As Variant)
    Me.Headers.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Replace the defined segment in the url with the provided value
'
' @param {String} Key
' @param {String} Value
' --------------------------------------------- '
Public Sub AddUrlSegment(Segment As String, Value As Variant)
    Me.UrlSegments.Item(Segment) = Value
End Sub

''
' Add querystring parameter to request
'
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddQuerystringParam(Key As String, Value As Variant)
    Me.QuerystringParams.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add cookie to request
'
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddCookie(Key As String, Value As Variant)
    Me.Cookies.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add parameter to body of request
'
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddBodyParameter(Key As String, Value As Variant)
    If IsEmpty(pBody) Then
        Set pBody = New Dictionary
    ElseIf Not TypeOf pBody Is Dictionary Then
        Err.Raise vbObjectError + 1, "WebRequest", "The existing body is not a Dictionary. Adding body parameters can only be used with Dictionaries"
    End If
    
    If VBA.IsObject(Value) Then
        Set pBody(Key) = Value
    Else
        pBody(Key) = Value
    End If
End Sub

''
' Prepare request for execution
' --------------------------------------------- '
Public Sub Prepare()
    ' Add general headers to request
    Me.AddHeader "User-Agent", WebUserAgent
    Me.AddHeader "Content-Type", Me.ContentType
    Me.AddHeader "Accept", Me.Accept
    Me.AddHeader "Content-Length", CStr(Me.ContentLength)
End Sub

''
' Clone request
'
' @return {WebRequest}
' --------------------------------------------- '
Public Function Clone() As WebRequest
    Set Clone = New WebRequest
    
    Clone.Resource = Me.Resource
    Clone.Method = Me.Method
    Clone.Accept = Me.Accept
    Clone.ContentType = Me.ContentType
    Clone.RequestFormat = Me.RequestFormat
    Clone.ResponseFormat = Me.ResponseFormat
    Clone.CustomRequestFormat = Me.CustomRequestFormat
    Clone.CustomResponseFormat = Me.CustomResponseFormat
    
    Set Clone.Headers = WebHelpers.CloneCollection(Me.Headers)
    Set Clone.QuerystringParams = WebHelpers.CloneCollection(Me.QuerystringParams)
    Set Clone.UrlSegments = WebHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = WebHelpers.CloneCollection(Me.Cookies)
    
    If VBA.IsObject(pBody) Then
        Set Clone.Body = pBody
    Else
        Clone.Body = pBody
    End If
End Function

''
' Create from options
'
' @param {Dictionary} Options
' - Headers {Collection}
' - Cookies {Collection}
' - QuerystringParams {Collection}
' - UrlSegments {DIctionary}
' --------------------------------------------- '
Public Sub CreateFromOptions(Options As Dictionary)
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set Me.Headers = Options("Headers")
        End If
        If Options.Exists("Cookies") Then
            Set Me.Cookies = Options("Cookies")
        End If
        If Options.Exists("QuerystringParams") Then
            Set Me.QuerystringParams = Options("QuerystringParams")
        End If
        If Options.Exists("UrlSegments") Then
            Set Me.UrlSegments = Options("UrlSegments")
        End If
    End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    ' Set default values
    Me.RequestFormat = WebFormat.Json
    Me.ResponseFormat = WebFormat.Json
    Set Me.Headers = New Collection
    Set Me.QuerystringParams = New Collection
    Set Me.UrlSegments = New Dictionary
    Set Me.Cookies = New Collection
    Me.ContentLength = -1
End Sub
