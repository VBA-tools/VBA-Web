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
' WebRequest v4.0.0-rc.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Create a request for use with a WebClient
'
' Errors:
' 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private web_pRequestFormat As WebFormat
Private web_pResponseFormat As WebFormat
Private web_pCustomRequestFormat As String
Private web_pCustomResponseFormat As String
Private web_pBody As Variant
Private web_pConvertedBody As String
Private web_pContentType As String
Private web_pAccept As String
Private web_pContentLength As Long
Private web_pId As String

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
    RequestFormat = web_pRequestFormat
End Property
Public Property Let RequestFormat(Value As WebFormat)
    web_pRequestFormat = Value
    web_pConvertedBody = ""
End Property

Public Property Get ResponseFormat() As WebFormat
    ResponseFormat = web_pResponseFormat
End Property
Public Property Let ResponseFormat(Value As WebFormat)
    web_pResponseFormat = Value
    web_pConvertedBody = ""
End Property

Public Property Get CustomRequestFormat() As String
    CustomRequestFormat = web_pCustomRequestFormat
End Property
Public Property Let CustomRequestFormat(Value As String)
    web_pCustomRequestFormat = Value
    web_pConvertedBody = ""
    
    If Value <> "" Then
        web_pRequestFormat = WebFormat.Custom
    End If
End Property

Public Property Get CustomResponseFormat() As String
    CustomResponseFormat = web_pCustomResponseFormat
End Property
Public Property Let CustomResponseFormat(Value As String)
    web_pCustomResponseFormat = Value
    web_pConvertedBody = ""
    
    If Value <> "" Then
        ResponseFormat = WebFormat.Custom
    End If
End Property

Public Property Get ContentType() As String
    If web_pContentType <> "" Then
        ContentType = web_pContentType
    Else
        ContentType = WebHelpers.FormatToMediaType(Me.RequestFormat, Me.CustomRequestFormat)
    End If
End Property
Public Property Let ContentType(Value As String)
    web_pContentType = Value
End Property

Public Property Get Accept() As String
    If web_pAccept <> "" Then
        Accept = web_pAccept
    Else
        Accept = WebHelpers.FormatToMediaType(Me.ResponseFormat, Me.CustomResponseFormat)
    End If
End Property
Public Property Let Accept(Value As String)
    web_pAccept = Value
End Property

Public Property Get ContentLength() As Long
    If web_pContentLength >= 0 Then
        ContentLength = web_pContentLength
    Else
        ContentLength = Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(Value As Long)
    web_pContentLength = Value
End Property

Public Property Get Body() As Variant
    If Not IsEmpty(web_pBody) Then
        If VarType(web_pBody) = vbString Then
            Body = web_pBody
        ElseIf web_pConvertedBody = "" Then
            ' Convert body and cache
            Body = WebHelpers.ConvertToFormat(web_pBody, Me.RequestFormat, Me.CustomRequestFormat)
            web_pConvertedBody = Body
        Else
            Body = web_pConvertedBody
        End If
    End If
End Property
Public Property Let Body(Value As Variant)
    web_pConvertedBody = ""
    web_pBody = Value
End Property
Public Property Set Body(Value As Variant)
    web_pConvertedBody = ""
    Set web_pBody = Value
End Property

Public Property Get FormattedResource() As String
    FormattedResource = Me.Resource
    
    ' Replace url segments
    Dim web_Segment As Variant
    For Each web_Segment In Me.UrlSegments.Keys
        FormattedResource = VBA.Replace(FormattedResource, "{" & web_Segment & "}", Me.UrlSegments(web_Segment))
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
    If web_pId = "" Then: web_pId = WebHelpers.CreateNonce
    Id = web_pId
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
    If VBA.IsEmpty(web_pBody) Then
        Set web_pBody = New Dictionary
    ElseIf Not TypeOf web_pBody Is Dictionary Then
        WebHelpers.LogError "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)", _
            "WebRequest.AddBodyParameter", 11020 + vbObjectError
        Err.Raise 11020 + vbObjectError, "WebRequest.AddBodyParameter", _
            "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)"
    End If
    
    If VBA.IsObject(Value) Then
        Set web_pBody(Key) = Value
    Else
        web_pBody(Key) = Value
    End If
End Sub

''
' Prepare request for execution
' --------------------------------------------- '
Public Sub Prepare()
    ' Add/replace general headers for request
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, "User-Agent", WebUserAgent
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, "Content-Type", Me.ContentType
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, "Accept", Me.Accept
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, "Content-Length", VBA.CStr(Me.ContentLength)
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
    
    If VBA.IsObject(web_pBody) Then
        Set Clone.Body = web_pBody
    Else
        Clone.Body = web_pBody
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
