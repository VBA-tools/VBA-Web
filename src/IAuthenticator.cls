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
' IAuthenticator v3.0.5
' (c) Tim Hall - https://github.com/timhall/Excel-REST
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
' @param {RestClient} Client The client that is about to execute the request
' @param {RestRequest} Request The request about to be executed
' --------------------------------------------- '

Public Sub BeforeExecute(ByVal Client As RestClient, ByRef Request As RestRequest)
    ' Add headers, cookies, etc
End Sub

''
' Hook for taking action after request has been executed
'
' @param {RestClient} Client The client that executed request
' @param {RestRequest} Request The request that was just executed
' @param {RestResponse} Response to request
' --------------------------------------------- '

Public Sub AfterExecute(ByVal Client As RestClient, ByVal Request As RestRequest, ByRef Response As RestResponse)
    ' Handle 401 Unauthorized or other issues
End Sub

''
' Hook for overriding standard http open (used for HTTP Basic)
'
' @param {MSXML2.IXMLHTTPRequest} http
' @parma {RestClient} Client The client that is about to open request
' @param {RestRequest} Request The request about to be opened
' @param {String} BaseUrl
' @param {Boolean} [useAsync=False]
' --------------------------------------------- '

Public Sub HttpOpen(ByRef Http As Object, ByVal Client As RestClient, ByRef Request As RestRequest, BaseUrl As String, Optional UseAsync As Boolean = False)
    ' Modify http open behavior
End Sub
