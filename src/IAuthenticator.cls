VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' IAuthenticator v1.0.3
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Interface for creating authenticators for rest client
'
' @dependencies
'   Microsoft XML, v3+
' @author tim.hall.engr@gmail.com
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
' @param {RestRequest} request The request about to be executed
' --------------------------------------------- '

Public Sub BeforeExecute(ByRef request As RestRequest)
    ' ...
    ' Add header
    ' etc.
End Sub

''
' Hook for overriding standard http open (used for HTTP Basic)
'
' @param {MSXML2.IXMLHTTPRequest} http
' @param {RestRequest} request The request about to be opened
' @param {String} BaseUrl
' @param {Boolean} [useAsync=False]
' --------------------------------------------- '

Public Sub HttpOpen(ByRef http As MSXML2.IXMLHTTPRequest, ByRef request As RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Use standard http open by default
    Call http.Open(request.MethodName(), request.FullUrl(BaseUrl), useAsync)
End Sub
