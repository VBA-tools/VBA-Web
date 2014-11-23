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
    Dim StatusCode As Long
    Dim StatusText As String
    Dim Headers As String
    Dim Body As Variant
    Dim ResponseText As String
    
    Dim Lines() As String
    Lines = Split(Result, vbCrLf)
    
    ' Extract status code and text from status line
    Dim StatusLine As String
    Dim StatusLineParts() As String
    StatusLine = Lines(0)
    StatusLineParts = Split(StatusLine)
    
    Me.StatusCode = CLng(StatusLineParts(1))
    Me.StatusDescription = Mid$(StatusLine, InStr(1, StatusLine, StatusCode) + 4)
    
    ' Find blank line before body
    Dim Line As Variant
    Dim BlankLineIndex
    BlankLineIndex = 0
    For Each Line In Lines
        If Trim(Line) = "" Then
            Exit For
        End If
        BlankLineIndex = BlankLineIndex + 1
    Next Line
    
    ' Extract body and headers strings
    Dim HeaderLines() As String
    Dim BodyLines() As String
    Dim ReadIndex As Long
    Dim WriteIndex As Long
    
    ReDim HeaderLines(0 To BlankLineIndex - 2)
    ReDim BodyLines(0 To UBound(Lines) - BlankLineIndex - 1)
    
    WriteIndex = 0
    For ReadIndex = 1 To BlankLineIndex - 1
        HeaderLines(WriteIndex) = Lines(ReadIndex)
        WriteIndex = WriteIndex + 1
    Next ReadIndex
    
    WriteIndex = 0
    For ReadIndex = BlankLineIndex + 1 To UBound(Lines)
        BodyLines(WriteIndex) = Lines(ReadIndex)
        WriteIndex = WriteIndex + 1
    Next ReadIndex
    
    Me.Content = VBA.Join$(BodyLines, vbCrLf)
    
    ' TODO: This needs to support UTF-8 or Unicode (match WinHttpRequest)
    ' Me.Body = WebHelpers.StringToUTF8Bytes(ResponseText)
    
    ' Convert content to data by format
    If Request.ResponseFormat <> WebFormat.PlainText Then
        Set CreateResponseFromCURL.Data = _
            WebHelpers.ParseByFormat(Me.Content, Request.ResponseFormat, Request.CustomResponseFormat, Me.Body)
    End If
    
    ' Extract headers
    Set Me.Headers = ExtractHeaders(VBA.Join$(HeaderLines, vbCrLf))
    
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
    Dim Headers As New Collection
    Dim Header As Dictionary
    Dim Multiline As Boolean
    Dim Key As String
    Dim Value As String
    
    Dim Lines As Variant
    Lines = Split(ResponseHeaders, vbCrLf)
    
    Dim i As Integer
    For i = LBound(Lines) To (UBound(Lines) + 1)
        If i > UBound(Lines) Then
            Headers.Add Header
        ElseIf Lines(i) <> "" Then
            If InStr(1, Lines(i), ":") = 0 And Not Header Is Nothing Then
                ' Assume part of multi-line header
                Multiline = True
            ElseIf Multiline Then
                ' Close out multi-line string
                Multiline = False
                Headers.Add Header
            ElseIf Not Header Is Nothing Then
                Headers.Add Header
            End If
            
            If Not Multiline Then
                Set Header = CreateKeyValue( _
                    Key:=Trim(Mid$(Lines(i), 1, InStr(1, Lines(i), ":") - 1)), _
                    Value:=Trim(Mid$(Lines(i), InStr(1, Lines(i), ":") + 1, Len(Lines(i)))) _
                )
            Else
                Header("Value") = Header("Value") & vbCrLf & Lines(i)
            End If
        End If
    Next i
    
    Set ExtractHeaders = Headers
End Function

''
' Extract cookies from response headers
'
' @param {Collection} Headers
' @return {Collection} Cookies
' @internal
' --------------------------------------------- '
Public Function ExtractCookies(Headers As Collection) As Collection
    Dim Cookies As New Collection
    Dim Cookie As String
    Dim Key As String
    Dim Value As String
    Dim Header As Dictionary
    
    For Each Header In Headers
        If Header("Key") = "Set-Cookie" Then
            Cookie = Header("Value")
            Key = Mid$(Cookie, 1, InStr(1, Cookie, "=") - 1)
            Value = Mid$(Cookie, InStr(1, Cookie, "=") + 1, Len(Cookie))
            
            If InStr(1, Value, ";") Then
                Value = Mid$(Value, 1, InStr(1, Value, ";") - 1)
            End If
            
            Cookies.Add CreateKeyValue(Key, UrlDecode(Value))
        End If
    Next Header
    
    Set ExtractCookies = Cookies
End Function

Private Sub Class_Initialize()
    Body = Array()
    Set Headers = New Collection
    Set Cookies = New Collection
End Sub
