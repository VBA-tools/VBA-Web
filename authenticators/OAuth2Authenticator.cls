VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuth2Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IAuthenticator
''
' OAuth2 Authenticator v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Utilize OAuth2 authentication
' (Currently using client credentials flow only)
'
' @dependencies
'   Microsoft XML, v3+
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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

Private Sub IAuthenticator_BeforeExecute(request As RestRequest)
    On Error GoTo ErrorHandling
    If (Me.Token = "" Or Not Me.CacheToken) And (Me.TokenUrl <> "" And Me.TokenKey <> "") Then
        ' Get new token
        Dim http As MSXML2.ServerXMLHTTP
        Set http = New MSXML2.ServerXMLHTTP
             
        Call http.Open("POST", CreateTokenRequest, False)
        Call http.send
        
        If http.status <> 200 Then
            ' Error getting OAuth2 token
            Err.Raise vbObjectError + http.status, _
                description:="Failed to retrieve OAuth2 Token - " & http.status & ": " & http.responseText
        End If
        
        Dim response As Dictionary
        Set response = RestHelpers.ParseJSON(http.responseText)
        If Not response Is Nothing Then
            If response.exists(Me.TokenKey) Then
                Me.Token = response(Me.TokenKey)
            End If
            
            ' (Salesforce specific, but shouldn't affect any other OAuth2 clients)
            If response.exists("instance_url") Then
                request.BaseUrl = response("instance_url")
            End If
        Else
            Err.Raise vbObjectError + 2, _
                description:="Failed to read OAuth2 Token"
        End If
    End If
    Call request.AddHeader("Authorization", CreateHeader())
    
ErrorHandling:
    
    If Not http Is Nothing Then Set http = Nothing
    If Not response Is Nothing Then Set response = Nothing
    
    If Err.Number > 0 Then
        ' Rethrow error
        Err.Raise Err.Number, description:=Err.description
    End If
End Sub

Private Sub IAuthenticator_HttpOpen(http As MSXML2.IXMLHTTPRequest, request As RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Perform standard http open
    Call http.Open(request.MethodName(), request.FullUrl(BaseUrl), useAsync)
End Sub

Private Function CreateHeader() As String
    ' Create standard OAuth2 header
    CreateHeader = "OAuth " & Me.Token
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

