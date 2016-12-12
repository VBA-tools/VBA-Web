VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OPSAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' OPSAuthenticator v3.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' OAuth2 client credentials authenticator for OPS
' http://documents.epo.org/projects/babylon/eponet.nsf/0/7AF8F1D2B36F3056C1257C04002E0AD6/$File/OPS_v3.1_documentation_version_1.2.14_en.pdf
'
' @class OPSAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

Public ConsumerKey As String
Public ConsumerSecret As String

Public Token As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup authenticator
'
' @param {String} ConsumerKey
' @param {String} ConsumerSecret
''
Public Sub Setup(ConsumerKey As String, ConsumerSecret As String)
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    If Me.Token = "" Then
        Me.Token = Me.GetToken(Client)
    End If
    
    Request.SetHeader "Authorization", "Bearer " & Me.Token
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
    ' e.g. Update option, headers, etc.
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
End Sub

Public Function GetToken(Client As WebClient) As String
    Dim TokenClient As WebClient
    Dim TokenAuth As New HttpBasicAuthenticator
    Dim TokenRequest As New WebRequest
    Dim Encoded As String
    Dim TokenResponse As WebResponse
    
    ' Clone client to avoid accidental interactions
    Set TokenClient = Client.Clone
    
    ' OPS uses Basic Auth to pass consumer key and secret
    TokenAuth.Setup Me.ConsumerKey, Me.ConsumerSecret
    Set TokenClient.Authenticator = TokenAuth
    
    ' Setup request according to docs
    TokenRequest.Resource = "auth/accesstoken"
    TokenRequest.RequestFormat = WebFormat.FormUrlEncoded
    TokenRequest.ResponseFormat = WebFormat.Json
    TokenRequest.Method = WebMethod.HttpPost
    TokenRequest.AddBodyParameter "grant_type", "client_credentials"
    
    Set TokenResponse = TokenClient.Execute(TokenRequest)
    If TokenResponse.StatusCode = WebStatusCode.Ok Then
        GetToken = TokenResponse.Data("access_token")
    Else
        WebHelpers.LogError "Failed to load token, " & TokenResponse.StatusCode & ": " & TokenResponse.Content
    End If
End Function
