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

Private pHeaders As Collection
Private pQuerystringParams As Collection
Private pUrlSegments As Dictionary
Private pCookies As Collection
Private pBody As Variant
Private pContentType As String
Private pAccept As String
Private pContentLength As Long
Private pId As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Resource As String
Public Method As WebMethod
Public RequestFormat As WebFormat
Public ResponseFormat As WebFormat

Public Property Get Format() As WebFormat
    Format = RequestFormat
End Property
Public Property Let Format(Value As WebFormat)
    Me.RequestFormat = Value
    Me.ResponseFormat = Value
End Property

Public Property Get Headers() As Collection
    If pHeaders Is Nothing Then: Set pHeaders = New Collection
    Set Headers = pHeaders
End Property
Public Property Set Headers(Value As Collection)
    Set pHeaders = Value
End Property

Public Property Get QuerystringParams() As Collection
    If pQuerystringParams Is Nothing Then: Set pQuerystringParams = New Collection
    Set QuerystringParams = pQuerystringParams
End Property
Public Property Set QuerystringParams(Value As Collection)
    Set pQuerystringParams = Value
End Property

Public Property Get UrlSegments() As Dictionary
    If pUrlSegments Is Nothing Then: Set pUrlSegments = New Dictionary
    Set UrlSegments = pUrlSegments
End Property
Public Property Set UrlSegments(Value As Dictionary)
    Set pUrlSegments = Value
End Property

Public Property Get Cookies() As Collection
    If pCookies Is Nothing Then: Set pCookies = New Collection
    Set Cookies = pCookies
End Property
Public Property Set Cookies(Value As Collection)
    Set pCookies = Value
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
        Else
            If Me.RequestFormat <> WebFormat.plaintext Then
                Body = RestHelpers.ConvertToFormat(pBody, Me.RequestFormat)
            End If
        End If
    End If
End Property
Public Property Let Body(Value As Variant)
    pBody = Value
End Property
Public Property Set Body(Value As Variant)
    Set pBody = Value
End Property

Public Property Get FormattedResource() As String
    Dim Segment As Variant
    
    FormattedResource = Me.Resource
    
    ' Replace url segments
    For Each Segment In Me.UrlSegments.Keys
        FormattedResource = Replace(FormattedResource, "{" & Segment & "}", Me.UrlSegments(Segment))
    Next Segment
    
    ' Add querystring
    If Me.QuerystringParams.Count > 0 Then
        If VBA.InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If
        
        FormattedResource = FormattedResource & RestHelpers.ConvertToUrlEncoded(Me.QuerystringParams)
    End If
End Property

Public Property Get Id() As String
    If pId = "" Then
        pId = RestHelpers.CreateNonce
    End If
    
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
    Me.Headers.Add RestHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Replace the defined segment in the url with the provided value
'
' @param {String} Key
' @param {String} Value
' --------------------------------------------- '
Public Sub AddUrlSegment(Segment As String, Value As Variant)
    Me.UrlSegments(Segment) = Value
End Sub

''
' Add querystring parameter to request
'
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddQuerystringParam(Key As String, Value As Variant)
    Me.QuerystringParams.Add RestHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add cookie to request
'
' @param {String} Key
' @param {Variant} Value
' --------------------------------------------- '
Public Sub AddCookie(Key As String, Value As Variant)
    Me.Cookies.Add RestHelpers.CreateKeyValue(Key, Value)
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
        Err.Raise vbObjectError + 1, "RestRequest", "The existing body is not a Dictionary. Adding body parameters can only be used with Dictionaries"
    End If
    
    If VBA.IsObject(Value) Then
        Set pBody(Key) = Value
    Else
        pBody(Key) = Value
    End If
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
    
    Set Clone.Headers = RestHelpers.CloneCollection(Me.Headers)
    Set Clone.QuerystringParams = RestHelpers.CloneCollection(Me.QuerystringParams)
    Set Clone.UrlSegments = RestHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = RestHelpers.CloneCollection(Me.Cookies)
    
    If VBA.IsObject(pBody) Then
        Set Clone.Body = pBody
    Else
        Clone.Body = pBody
    End If
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    ' Set default values
    Me.RequestFormat = WebFormat.json
    Me.ResponseFormat = WebFormat.json
    Me.ContentLength = -1
End Sub
