VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' IWebAuthenticator v4.0.0-beta.5
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Interface for creating authenticators for rest client
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' ============================================= '
' Public Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param {WebRequest} Request The request about to be executed
' --------------------------------------------- '

Public Sub BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' Add headers, cookies, etc
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param {WebResponse} Response to request
' --------------------------------------------- '

Public Sub AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' Handle 401 Unauthorized or other issues
End Sub
