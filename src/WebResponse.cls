VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebResponse v4.0.0-beta.3
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Wrapper for http responses
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public StatusCode As WebStatusCode
Public StatusDescription As String
Public Content As String
Public Data As Object
Public Body As Variant
Public Headers As Collection
Public Cookies As Collection

' ============================================= '
' Public Methods
' ============================================= '

''
' Update response with given response values
' (Useful with ByRef)
'
' @param {WebResponse} Updated
' --------------------------------------------- '
Public Sub Update(Updated As WebResponse)
    Me.StatusCode = Updated.StatusCode
    Me.StatusDescription = Updated.StatusDescription
    Me.Content = Updated.Content
    Me.Body = Updated.Body
    Set Me.Headers = Updated.Headers
    Set Me.Cookies = Updated.Cookies
    Set Me.Data = Updated.Data
End Sub

''
' Create response from http
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {WinHttpRequest} Http
' --------------------------------------------- '
Public Sub CreateFromHttp(Client As WebClient, Request As WebRequest, Http As Object)
    Me.StatusCode = Http.Status
    Me.StatusDescription = Http.StatusText
    Me.Body = Http.ResponseBody
    Me.Content = Http.ResponseText
    
    ' Convert content to data by format
    If Request.ResponseFormat <> WebFormat.PlainText Then
        Set Me.Data = _
            WebHelpers.ParseByFormat(Me.Content, Request.ResponseFormat, Request.CustomResponseFormat, Me.Body)
    End If
    
    ' Extract headers
    Set Me.Headers = ExtractHeaders(Http.getAllResponseHeaders)
    
    ' Extract cookies
    Set Me.Cookies = ExtractCookies(Me.Headers)
End Sub

''
' Create response from cURL
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {String} Result
' --------------------------------------------- '
Public Sub CreateFromCURL(Client As WebClient, Request As WebRequest, Result As String)
    Dim web_StatusCode As Long
    Dim web_StatusText As String
    Dim web_Headers As String
    Dim web_Body As Variant
    Dim web_ResponseText As String
    
    Dim web_Lines() As String
    web_Lines = Split(Result, vbCrLf)
    
    ' Extract status code and text from status line
    Dim web_StatusLine As String
    Dim web_StatusLineParts() As String
    web_StatusLine = web_Lines(0)
    web_StatusLineParts = Split(web_StatusLine)
    
    Me.StatusCode = CLng(web_StatusLineParts(1))
    Me.StatusDescription = Mid$(web_StatusLine, InStr(1, web_StatusLine, web_StatusCode) + 4)
    
    ' Find blank line before body
    Dim web_Line As Variant
    Dim web_BlankLineIndex As Long
    web_BlankLineIndex = 0
    For Each web_Line In web_Lines
        If Trim(web_Line) = "" Then
            Exit For
        End If
        web_BlankLineIndex = web_BlankLineIndex + 1
    Next web_Line
    
    ' Extract body and headers strings
    Dim web_HeaderLines() As String
    Dim web_BodyLines() As String
    Dim web_ReadIndex As Long
    Dim web_WriteIndex As Long
    
    ReDim web_HeaderLines(0 To web_BlankLineIndex - 2)
    ReDim web_BodyLines(0 To UBound(web_Lines) - web_BlankLineIndex - 1)
    
    web_WriteIndex = 0
    For web_ReadIndex = 1 To web_BlankLineIndex - 1
        web_HeaderLines(web_WriteIndex) = web_Lines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex
    
    web_WriteIndex = 0
    For web_ReadIndex = web_BlankLineIndex + 1 To UBound(web_Lines)
        web_BodyLines(web_WriteIndex) = web_Lines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex
    
    Me.Content = VBA.Join$(web_BodyLines, vbCrLf)
    Me.Body = WebHelpers.StringToANSIBytes(Me.Content)
    
    ' Convert content to data by format
    If Request.ResponseFormat <> WebFormat.PlainText Then
        Set Me.Data = _
            WebHelpers.ParseByFormat(Me.Content, Request.ResponseFormat, Request.CustomResponseFormat, Me.Body)
    End If
    
    ' Extract headers
    Set Me.Headers = ExtractHeaders(VBA.Join$(web_HeaderLines, vbCrLf))
    
    ' Extract cookies
    Set Me.Cookies = ExtractCookies(Me.Headers)
End Sub

''
' Extract headers from response headers
'
' @param {String} ResponseHeaders
' @return {Collection} Headers
' @internal
' --------------------------------------------- '
Public Function ExtractHeaders(ResponseHeaders As String) As Collection
    Dim web_Headers As New Collection
    Dim web_Header As Dictionary
    Dim web_Multiline As Boolean
    Dim web_Key As String
    Dim web_Value As String
    
    Dim web_Lines As Variant
    web_Lines = Split(ResponseHeaders, vbCrLf)
    
    Dim web_i As Integer
    For web_i = LBound(web_Lines) To (UBound(web_Lines) + 1)
        If web_i > UBound(web_Lines) Then
            web_Headers.Add web_Header
        ElseIf web_Lines(web_i) <> "" Then
            If InStr(1, web_Lines(web_i), ":") = 0 And Not web_Header Is Nothing Then
                ' Assume part of multi-line header
                web_Multiline = True
            ElseIf web_Multiline Then
                ' Close out multi-line string
                web_Multiline = False
                web_Headers.Add web_Header
            ElseIf Not web_Header Is Nothing Then
                web_Headers.Add web_Header
            End If
            
            If Not web_Multiline Then
                Set web_Header = CreateKeyValue( _
                    Key:=Trim(Mid$(web_Lines(web_i), 1, InStr(1, web_Lines(web_i), ":") - 1)), _
                    Value:=Trim(Mid$(web_Lines(web_i), InStr(1, web_Lines(web_i), ":") + 1, Len(web_Lines(web_i)))) _
                )
            Else
                web_Header("Value") = web_Header("Value") & vbCrLf & web_Lines(web_i)
            End If
        End If
    Next web_i
    
    Set ExtractHeaders = web_Headers
End Function

''
' Extract cookies from response headers
'
' @param {Collection} Headers
' @return {Collection} Cookies
' @internal
' --------------------------------------------- '
Public Function ExtractCookies(Headers As Collection) As Collection
    Dim web_Cookies As New Collection
    Dim web_Cookie As String
    Dim web_Key As String
    Dim web_Value As String
    Dim web_Header As Dictionary
    
    For Each web_Header In Headers
        If web_Header("Key") = "Set-Cookie" Then
            web_Cookie = web_Header("Value")
            web_Key = Mid$(web_Cookie, 1, InStr(1, web_Cookie, "=") - 1)
            web_Value = Mid$(web_Cookie, InStr(1, web_Cookie, "=") + 1, Len(web_Cookie))
            
            If InStr(1, web_Value, ";") > 0 Then
                web_Value = Mid$(web_Value, 1, InStr(1, web_Value, ";") - 1)
            End If
            
            web_Cookies.Add CreateKeyValue(web_Key, UrlDecode(web_Value))
        End If
    Next web_Header
    
    Set ExtractCookies = web_Cookies
End Function

Private Sub Class_Initialize()
    Set Headers = New Collection
    Set Cookies = New Collection
End Sub
