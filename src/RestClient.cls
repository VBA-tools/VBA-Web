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
' RestClient v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Interact with REST web services from Excel
'
' @dependencies
'   Microsoft Scripting Runtime
'   Microsoft XML, v3+
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const UserAgent As String = "Excel Client v1.0.0 (https://github.com/timhall/Excel-REST)"
Private Const Timeoutms As Integer = 5000


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
' @param {RestRequest} request The request to execute
' @return {RestResponse} Wrapper of server response for request
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
    response.StatusCode = http.status
    response.StatusDescription = http.statusText
    response.Content = http.responseText
    Select Case request.Format
    Case Else
        Set response.Data = RestHelpers.ParseJSON(response.Content)
    End Select
    Set Execute = response
    
ErrorHandling:

    If Not http Is Nothing Then Set http = Nothing
    If Not response Is Nothing Then Set response = Nothing
    
    If Err.Number > 0 Then
        ' Rethrow error
        Err.Raise Err.Number, description:=Err.description
    End If
    
End Function

''
' Execute the specified request asynchronously
'
' @param {RestRequest} request The request to execute
' @param {String} callback Name of function to call when request completes (specify "" if none)
' @param {Variant} [callbackArgs] Variable array of arguments that get passed directly to callback function
' @return {Boolean} Status of initiating request
' --------------------------------------------- '

Public Function ExecuteAsync(request As RestRequest, Callback As String, Optional ByVal CallbackArgs As Variant) As Boolean
    Dim response As New RestResponse
    Dim http As MSXML2.ServerXMLHTTP
    
    On Error GoTo ErrorHandling
    
    ' Setup the request
    Set http = New MSXML2.ServerXMLHTTP
    Call HttpSetup(http, request, True)
    request.Callback = Callback
    request.CallbackArgs = CallbackArgs
    
    ' Send the request
    Call http.send(request.Body)
    
    ' Clean up and return
    ' Set http = Nothing
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    If Not http Is Nothing Then Set http = Nothing
    If Not response Is Nothing Then Set response = Nothing
    
    If Err.Number > 0 Then
        ' Rethrow error
        Err.Raise Err.Number, description:=Err.description
    End If
    
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub HttpSetup(ByRef http As MSXML2.ServerXMLHTTP, ByRef request As RestRequest, Optional useAsync As Boolean = False)
    ' Set timeouts
    http.setTimeouts Timeoutms, Timeoutms, Timeoutms, Timeoutms
    
    ' Add general headers to request
    Call request.AddHeader("User-Agent", UserAgent)
    Call request.AddHeader("Content-Type", request.ContentType())
    
    ' If async, pass http to request and setup onreadystatechange
    If useAsync Then
        request.HttpRequest = http
        http.onreadystatechange = request
    End If
    
    ' Before execute and http open hooks for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute request
        Me.Authenticator.HttpOpen http, request, Me.BaseUrl, useAsync
    Else
        ' Nothing hooked in so open http object
        Call http.Open(request.MethodName(), request.FullUrl(Me.BaseUrl), useAsync)
    End If
    
    ' Set request headers
    Dim headerKey As Variant
    For Each headerKey In request.Headers.keys()
        Call http.setRequestHeader(headerKey, request.Headers(headerKey))
    Next headerKey
End Sub
