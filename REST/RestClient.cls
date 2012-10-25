VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' RestClient v0.2.1
' (c) Tim Hall - https://github.com/timhall/ExcelHelpers
'
' Interact with REST web services from Excel
'
' @dependencies
'   Microsoft Scripting Runtime
'   Microsoft XML, v3+
'   JSONLib (http://code.google.com/p/vba-json/)
' @author tim.hall.engr@gmail.com
' @version 0.2.1
' @date 20120826
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '
Private Const UserAgent As String = "Excel Client v0.2.1 (tim.hall.engr@gmail.com)"

' --------------------------------------------- '
' Types
' --------------------------------------------- '


' --------------------------------------------- '
' Properties
' --------------------------------------------- '
Public BaseUrl As String
Public Authenticator As IAuthenticator

' ============================================= '
' Public Methods
' ============================================= '

''
' Execute the specified request
'
' @param request    The request to execute
' @return           Server response for request
' --------------------------------------------- '
Public Function Execute(request As RestRequest) As RestResponse
    Dim response As New RestResponse
    Dim http As MSXML2.ServerXMLHTTP
    Dim headerKey As Variant
    
    On Error GoTo ErrorHandling
    Set http = New MSXML2.ServerXMLHTTP
    Call HttpSetup(http, request, False)
    
    ' Send the request
    Call http.send(request.Body)
    
    ' Handle response...
    response.StatusCode = http.Status
    response.StatusDescription = http.statusText
    response.Content = http.responseText
    Select Case request.Format
    Case Else
        Set response.Data = RestHelpers.ParseJSON(response.Content)
    End Select
    
    ' Clean up and return
    Set http = Nothing
    Set Execute = response
    Exit Function
    
ErrorHandling:

    If Not http Is Nothing Then Set http = Nothing
    ' Do something with error...
    
End Function

''
' Execute the specified request asynchronously
'
' @param request    The request to execute
' @return           Server response for request
' --------------------------------------------- '
Public Function ExecuteAsync(request As RestRequest, callback As String, Optional ByVal CallbackArgs As Variant) As Boolean
    Dim response As New RestResponse
    Dim http As MSXML2.ServerXMLHTTP
    
    ' On Error GoTo ErrorHandling
    
    ' Setup the request
    Set http = New MSXML2.ServerXMLHTTP
    Call HttpSetup(http, request, True)
    request.callback = callback
    request.CallbackArgs = CallbackArgs
    
    ' Send the request
    Call http.send(request.Body)
    
    ' Clean up and return
    ' Set http = Nothing
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    Set http = Nothing
    ' Do something with error...
    
End Function

' ============================================= '
' Private Methods
' ============================================= '
Private Sub HttpSetup(ByRef http As MSXML2.ServerXMLHTTP, ByRef request As RestRequest, Optional useAsync As Boolean = False)
    ' Set timeouts
    http.setTimeouts 5 * 1000, 5 * 1000, 5 * 1000, 5 * 1000
    
    ' Add general headers to request
    Call request.AddHeader("User-Agent", UserAgent)
    Call request.AddHeader("Content-Type", request.ContentType())
    
    If useAsync Then
        request.HttpRequest = http
        http.onreadystatechange = request
    End If
    
    ' Before execute and http open hooks for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute request
        Me.Authenticator.HttpOpen http, request, Me.BaseUrl, useAsync
    Else
        ' Create new http object and open
        Call http.Open(request.MethodName(), request.FullUrl(Me.BaseUrl), useAsync)
    End If
    
    ' Set headers
    Dim headerKey As Variant
    For Each headerKey In request.Headers.keys()
        Call http.setRequestHeader(headerKey, request.Headers(headerKey))
    Next headerKey
End Sub






