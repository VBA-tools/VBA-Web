VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuth2Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' OAuth2 Authenticator v2.0.6
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Utilize OAuth2 authentication
' (Currently using client credentials flow only)
'
' @implements: IWebAuthenticator v3.*
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public ClientSecret As String
Public Username As String
Public Password As String
Public TokenUrl As String
Public TokenKey As String
Public Token As String
Public CacheToken As Boolean

' ============================================= '
' Public Methods
' ============================================= '

''
' Helper for setting required parameters at once
'
' @param {String} ClientId
' @param {String} ClientSecret
' @param {String} Username
' @param {String} Password
' --------------------------------------------- '

Public Sub Setup(ClientId As String, ClientSecret As String, Username As String, Password As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
    Me.Username = Username
    Me.Password = Password
End Sub

''
' Setup token url that is used to request token
'
' @param {String} TokenUrl
' @param {String} [TokenKey="access_token"] The key used to retrieve the token from the response
' --------------------------------------------- '

Public Sub SetupTokenUrl(TokenUrl As String, Optional TokenKey As String = "access_token")
    Me.TokenUrl = TokenUrl
    Me.TokenKey = TokenKey
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
    On Error GoTo ErrorHandling
    If (Me.Token = "" Or Not Me.CacheToken) And (Me.TokenUrl <> "" And Me.TokenKey <> "") Then
        ' Get new token
        Dim Http As Object
        Set Http = CreateObject("MSXML2.ServerXMLHTTP")
             
        Http.Open "POST", CreateTokenRequest, False
        Http.Send
        
        If Http.Status <> 200 Then
            ' Error getting OAuth2 token
            WebHelpers.LogError "Token request failed: " & Http.Status & " - " & Http.ResponseText, "OAuth2Authenticator.BeforeExecute"
            Err.Raise vbObjectError + Http.Status, _
                Description:="Failed to retrieve OAuth2 Token - " & Http.Status & ": " & Http.ResponseText
        End If
        
        Dim Response As Dictionary
        Set Response = WebHelpers.ParseJSON(Http.ResponseText)
        If Not Response Is Nothing Then
            If Response.Exists(Me.TokenKey) Then
                Me.Token = Response(Me.TokenKey)
                WebHelpers.LogDebug "Received token: " & WebHelpers.Obfuscate(Me.Token), "OAuth2Authenticator.BeforeExecute"
            End If
            
            ' (Salesforce specific, but shouldn't affect any other OAuth2 clients)
            If Response.Exists("instance_url") Then
                Client.BaseUrl = Response("instance_url")
            End If
        Else
            WebHelpers.LogError "Failed to read OAuth2 Token: " & Http.ResponseText, "OAuth2Authenticator.BeforeExecute"
            Err.Raise vbObjectError + 2, _
                Description:="Failed to read OAuth2 Token"
        End If
    End If
    Call Request.AddHeader("Authorization", CreateHeader())
    
ErrorHandling:
    
    If Not Http Is Nothing Then Set Http = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Description:=Err.Description
    End If
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

Private Function CreateHeader() As String
    ' Create standard OAuth2 header
    CreateHeader = "Bearer " & Me.Token
End Function

Private Function CreateTokenRequest() As String
    ' Generate token request using client credentials flow
    ' This is currently setup using Salesforce's implementation, but may need to be generalized
    ' [Digging Deeper into OAuth 2.0 on Force.com](http://wiki.developerforce.com/page/Digging_Deeper_into_OAuth_2.0_at_Salesforce.com)
    CreateTokenRequest = Me.TokenUrl
    If InStr(1, CreateTokenRequest, "?") < 1 Then: CreateTokenRequest = CreateTokenRequest & "?"
    CreateTokenRequest = CreateTokenRequest & "&client_id=" & Me.ClientId
    CreateTokenRequest = CreateTokenRequest & "&client_secret=" & Me.ClientSecret
    CreateTokenRequest = CreateTokenRequest & "&username=" & Me.Username
    CreateTokenRequest = CreateTokenRequest & "&password=" & Me.Password
End Function
