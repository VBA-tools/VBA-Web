VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Spec Authenticator
'
' @implements: IAuthenticator v3.*
' @author: Tim Hall
' @license: MIT
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IAuthenticator
Option Explicit

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup authenticator
' --------------------------------------------- '

Public Sub Setup()
    ' Define any user-specific variables needed for authentication
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {RestClient} Client The client that is about to execute the request
' @param {RestRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IAuthenticator_BeforeExecute(ByVal Client As RestClient, ByRef Request As RestRequest)
    ' It should:
    ' Add parameter, header, cookie, content-type
    Request.AddBodyParameter "auth_parameter", "auth"
    Request.AddQuerystringParam "auth_query", "auth"
    Request.ContentType = "text/plain"
    Request.AddHeader "X-Custom-A", "auth"
    Request.AddCookie "auth_cookie", "auth"
End Sub

''
' Hook for taking action after request has been executed
'
' @param {RestClient} Client The client that executed request
' @param {RestRequest} Request The request that was just executed
' @param {RestResponse} Response to request
' --------------------------------------------- '

Private Sub IAuthenticator_AfterExecute(ByVal Client As RestClient, ByVal Request As RestRequest, ByRef Response As RestResponse)
    ' It should:
    ' Update response
    Response.StatusCode = 204
End Sub

