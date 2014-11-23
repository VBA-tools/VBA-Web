VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwitterAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Twitter Authenticator v2.0.6
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Custom IWebAuthenticator for application-only authentication in Twitter's V1.1 REST API
'
' - https://dev.twitter.com/docs/auth/application-only-auth
' - https://github.com/timhall/VBA-Web/wiki/Implementing-your-own-IWebAuthenticator
'
' @implements: IWebAuthenticator v3.*
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

Private pToken As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ConsumerKey As String
Public ConsumerSecret As String
Public CacheToken As Boolean

Public Property Get Token() As String
    On Error GoTo Cleanup

    If pToken = "" Or Not CacheToken Then
        Dim Client As New WebClient
        Dim Response As WebResponse
        
        ' Request a new token
        Client.BaseUrl = "https://api.twitter.com/"
        Set Response = Client.Execute(TokenRequest)
        
        ' Store token if successful, otherwise throw error
        If Response.StatusCode = Ok Then
            pToken = Response.Data("access_token")
            WebHelpers.LogDebug "Successfully received token: " & WebHelpers.Obfuscate(pToken), "TwitterAuthenticator.Token"
        Else
            WebHelpers.LogError "Failed to load token: " & Response.StatusCode & " - " & Response.Content, "TwitterAuthenticator.Token"
            Err.Raise vbObjectError + Response.StatusCode, _
                Description:="Failed to load Bearer Token: " & Response.StatusCode & " - " & Response.Content
        End If
    End If
    
    ' Return token
    Token = pToken
    
Cleanup:

    If Not Client Is Nothing Then Set Client = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then Err.Raise Err.Number, Description:=Err.Description
    
End Property
Public Property Let Token(Value As String)
    pToken = Value
End Property

' ============================================= '
' Public Methods
' ============================================= '

Public Sub Setup(ConsumerKey As String, ConsumerSecret As String)
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param {WebRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    Request.AddHeader "Authorization", "Bearer " & Me.Token
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param {WebResponse} Response to request
' --------------------------------------------- '

Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    
End Sub

Private Function TokenRequest() As WebRequest
    Set TokenRequest = New WebRequest
    TokenRequest.Resource = "oauth2/token"
    TokenRequest.Method = HttpPost
    
    ' Based on documentation:
    ' 1. Pass in Base64 encoded ConsumerKey and ConsumerSecret
    TokenRequest.AddHeader "Authorization", _
        "Basic " & WebHelpers.Base64Encode(Me.ConsumerKey & ":" & Me.ConsumerSecret)

    ' 2. Include grant_type in body
    TokenRequest.Body = "grant_type=client_credentials"

    ' 3. Set content-type as form-urlencoded
    TokenRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
End Function

Private Sub Class_Initialize()
    ' Cache token by default
    Me.CacheToken = True
End Sub
