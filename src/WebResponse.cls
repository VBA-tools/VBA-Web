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
' WebResponse v4.0.0-rc.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Wrapper for http/cURL responses
'
' Errors:
' 11030 / 80042b16 / -2147210474 - Error creating from http
' 11031 / 80042b17 / -2147210473 - Error creating from cURL
' 11032 / 80042b18 / -2147210472 - Error extracting headers
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
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {WinHttpRequest} Http
' --------------------------------------------- '
Public Sub CreateFromHttp(Client As WebClient, Request As WebRequest, Http As Object)
    On Error GoTo web_ErrorHandling

    Me.StatusCode = Http.Status
    Me.StatusDescription = Http.StatusText
    Me.Body = Http.ResponseBody
    Me.Content = Http.ResponseText
    
    ' Convert content to data by format
    If Request.ResponseFormat <> WebFormat.PlainText Then
        On Error Resume Next
        Set Me.Data = _
            WebHelpers.ParseByFormat(Me.Content, Request.ResponseFormat, Request.CustomResponseFormat, Me.Body)
        
        If Err.Number <> 0 Then
            WebHelpers.LogError Err.Description, Err.Source, Err.Number
            Err.Clear
        End If
        On Error GoTo web_ErrorHandling
    End If
    
    ' Extract headers
    Set Me.Headers = ExtractHeaders(Http.getAllResponseHeaders)
    
    ' Extract cookies
    Set Me.Cookies = ExtractCookies(Me.Headers)
    Exit Sub
    
web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from http" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
    
    WebHelpers.LogError web_ErrorDescription, "WebResponse.CreateFromHttp", 11030 + vbObjectError
    Err.Raise 11030 + vbObjectError, "WebResponse.CreateFromHttp", web_ErrorDescription
End Sub

''
' Create response from cURL
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {String} Result
' --------------------------------------------- '
Public Sub CreateFromCurl(Client As WebClient, Request As WebRequest, Result As String)
    On Error GoTo web_ErrorHandling

    Dim web_Lines() As String
    web_Lines = VBA.Split(Result, vbCrLf)
    
    ' Extract status code and text from status line
    Dim web_StatusLine As String
    Dim web_StatusLineParts() As String
    Dim web_i As Long
    web_StatusLine = web_Lines(0)
    web_StatusLineParts = VBA.Split(web_StatusLine)
    
    Me.StatusCode = VBA.CLng(web_StatusLineParts(1))
    For web_i = 2 To UBound(web_StatusLineParts)
        If web_i > 2 Then
            Me.StatusDescription = Me.StatusDescription & " "
        End If
        
        Me.StatusDescription = Me.StatusDescription & web_StatusLineParts(web_i)
    Next web_i
    
    ' Find blank line before body
    Dim web_Line As Variant
    Dim web_BlankLineIndex As Long
    web_BlankLineIndex = 0
    For Each web_Line In web_Lines
        If VBA.Trim(web_Line) = "" Then
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
    Me.Body = WebHelpers.StringToAnsiBytes(Me.Content)
    
    ' Convert content to data by format
    If Request.ResponseFormat <> WebFormat.PlainText Then
        On Error Resume Next
        Set Me.Data = _
            WebHelpers.ParseByFormat(Me.Content, Request.ResponseFormat, Request.CustomResponseFormat, Me.Body)
            
        If Err.Number <> 0 Then
            WebHelpers.LogError Err.Description, Err.Source, Err.Number
            Err.Clear
        End If
        On Error GoTo web_ErrorHandling
    End If
    
    ' Extract headers
    Set Me.Headers = ExtractHeaders(VBA.Join$(web_HeaderLines, vbCrLf))
    
    ' Extract cookies
    Set Me.Cookies = ExtractCookies(Me.Headers)
    Exit Sub
    
web_ErrorHandling:
    
    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from cURL" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
    
    WebHelpers.LogError web_ErrorDescription, "WebResponse.CreateFromCurl", 11031 + vbObjectError
    Err.Raise 11031 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Sub

''
' Extract headers from response headers
'
' @internal
' @param {String} ResponseHeaders
' @return {Collection} Headers
' --------------------------------------------- '
Public Function ExtractHeaders(ResponseHeaders As String) As Collection
    On Error GoTo web_ErrorHandling
    
    Dim web_Headers As New Collection
    Dim web_Header As Dictionary
    Dim web_Multiline As Boolean
    Dim web_Key As String
    Dim web_Value As String
    Dim web_ColonPosition As Long
    
    Dim web_Lines As Variant
    web_Lines = VBA.Split(ResponseHeaders, vbCrLf)
    
    Dim web_i As Integer
    For web_i = LBound(web_Lines) To (UBound(web_Lines) + 1)
        If web_i > UBound(web_Lines) Then
            web_Headers.Add web_Header
        ElseIf web_Lines(web_i) <> "" Then
            web_ColonPosition = VBA.InStr(1, web_Lines(web_i), ":")
            If web_ColonPosition = 0 And Not web_Header Is Nothing Then
                ' Assume part of multi-line header
                web_Multiline = True
            ElseIf web_Multiline Then
                ' Close out multi-line string
                web_Multiline = False
                web_Headers.Add web_Header
            ElseIf Not web_Header Is Nothing Then
                ' Add previous header
                web_Headers.Add web_Header
            End If
            
            If Not web_Multiline Then
                Set web_Header = WebHelpers.CreateKeyValue( _
                    Key:=VBA.Trim(VBA.Mid$(web_Lines(web_i), 1, web_ColonPosition - 1)), _
                    Value:=VBA.Trim(VBA.Mid$(web_Lines(web_i), web_ColonPosition + 1, VBA.Len(web_Lines(web_i)))) _
                )
            Else
                web_Header("Value") = web_Header("Value") & vbCrLf & web_Lines(web_i)
            End If
        End If
    Next web_i
    
    Set ExtractHeaders = web_Headers
    Exit Function
    
web_ErrorHandling:
    
    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while extracting headers" & vbNewLine & _
        Err.Number & _
        VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & _
        ": " & Err.Description
    
    WebHelpers.LogError web_ErrorDescription, "WebResponse.CreateFromCurl", 11032 + vbObjectError
    Err.Raise 11032 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Function

''
' Extract cookies from response headers
'
' @internal
' @param {Collection} Headers
' @return {Collection} Cookies
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
            If VBA.InStr(1, web_Cookie, "=") > 0 Then
                web_Key = VBA.Mid$(web_Cookie, 1, VBA.InStr(1, web_Cookie, "=") - 1)
                web_Value = VBA.Mid$(web_Cookie, VBA.InStr(1, web_Cookie, "=") + 1, VBA.Len(web_Cookie))
                
                If VBA.InStr(1, web_Value, ";") > 0 Then
                    web_Value = VBA.Mid$(web_Value, 1, VBA.InStr(1, web_Value, ";") - 1)
                End If
                
                web_Cookies.Add WebHelpers.CreateKeyValue(web_Key, WebHelpers.UrlDecode(web_Value))
            Else
                WebHelpers.LogDebug _
                    "WARNING - Unrecognized cookie format: " & web_Cookie, "WebResponse.ExtractCookies"
            End If
        End If
    Next web_Header
    
    Set ExtractCookies = web_Cookies
End Function

Private Sub Class_Initialize()
    Set Headers = New Collection
    Set Cookies = New Collection
End Sub
