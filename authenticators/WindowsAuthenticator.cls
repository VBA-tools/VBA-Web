VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowsAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WindowsAuthenticator v3.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Use Windows (NTLM) authentication
'
' @class WindowsAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const web_AutoLogonPolicy_Always = 0
Private Const web_AutoLogonPolicy_OnlyIfBypassProxy = 1
Private Const web_AutoLogonPolicy_Never = 2

' ============================================= '
' Public Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' e.g Add headers, cookies, etc.
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
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384107(v=vs.85).aspx
    Http.SetAutoLogonPolicy web_AutoLogonPolicy_Always
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' http://curl.haxx.se/docs/manpage.html#--ntlm
    Curl = Curl & " --ntlm"
End Sub

