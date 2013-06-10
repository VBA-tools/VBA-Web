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
' RestClient v1.0.1
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

Private Const UserAgent As String = "Excel Client v1.0.1 (https://github.com/timhall/Excel-REST)"
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

Public Function Execute(Request As RestRequest) As RestResponse
    Dim Response As New RestResponse
    Dim http As MSXML2.ServerXMLHTTP
    Dim headerKey As Variant
    
    On Error GoTo ErrorHandling
    Set http = New MSXML2.ServerXMLHTTP
    Call HttpSetup(http, Request, False)
    
    ' Send the request
    Call http.send(Request.Body)
    
    ' Handle response...
    Response.StatusCode = http.Status
    Response.StatusDescription = http.statusText
    Response.Content = http.responseText
    Select Case Request.Format
    Case Else
        Set Response.Data = RestHelpers.ParseJSON(Response.Content)
    End Select
    Set Execute = Response
    
ErrorHandling:

    If Not http Is Nothing Then Set http = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Description:=Err.Description
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

Public Function ExecuteAsync(Request As RestRequest, Callback As String, Optional ByVal CallbackArgs As Variant) As Boolean
    Dim Response As New RestResponse
    Dim http As MSXML2.ServerXMLHTTP
    
    On Error GoTo ErrorHandling
    
    ' Setup the request
    Set http = New MSXML2.ServerXMLHTTP
    Call HttpSetup(http, Request, True)
    Request.Callback = Callback
    Request.CallbackArgs = CallbackArgs
    
    ' Send the request
    Call http.send(Request.Body)
    
    ' Clean up and return
    ' Set http = Nothing
    ExecuteAsync = True
    Exit Function
    
ErrorHandling:

    If Not http Is Nothing Then Set http = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Description:=Err.Description
    End If
    
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub HttpSetup(ByRef http As MSXML2.ServerXMLHTTP, ByRef Request As RestRequest, Optional useAsync As Boolean = False)
    ' Set timeouts
    http.setTimeouts Timeoutms, Timeoutms, Timeoutms, Timeoutms
    
    ' Add general headers to request
    Call Request.AddHeader("User-Agent", UserAgent)
    Call Request.AddHeader("Content-Type", Request.ContentType())
    
    ' If async, pass http to request and setup onreadystatechange
    If useAsync Then
        Request.HttpRequest = http
        http.onreadystatechange = Request
    End If
    
    ' Before execute and http open hooks for authenticator
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Request
        Me.Authenticator.HttpOpen http, Request, Me.BaseUrl, useAsync
    Else
        ' Nothing hooked in so open http object
        Call http.Open(Request.MethodName(), Request.FullUrl(Me.BaseUrl), useAsync)
    End If
    
    ' Set request headers
    Dim headerKey As Variant
    For Each headerKey In Request.Headers.keys()
        Call http.setRequestHeader(headerKey, Request.Headers(headerKey))
    Next headerKey
End Sub
