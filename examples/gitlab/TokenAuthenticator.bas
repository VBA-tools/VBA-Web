''
' Http Token Authenticator v3.0.5
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for GitLab Token Authenticator
'
' @class TokenAuthenticator
' @implements IWebAuthenticator v4.*
' @author mkysoft@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const web_HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0
Private Const web_HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Header As String
Public value As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} Header
' @param {String} Value
''
Public Sub Setup(Header As String, value As String)
    Me.Header = Header
    Me.value = value
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    Request.SetHeader Me.Header, Me.value
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
End Sub

''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
    'Http.SetCredentials Me.Username, Me.Password, web_HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' e.g. Add flags to cURL
    'Curl = Curl & " --basic --user " & WebHelpers.PrepareTextForShell(Me.Username) & ":" & WebHelpers.PrepareTextForShell(Me.Password)
End Sub
