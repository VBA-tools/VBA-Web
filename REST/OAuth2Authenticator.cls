VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EH_OAuth2Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements EH_IAuthenticator
''
' OAuth2 Authenticator v0.2.1
' (c) Tim Hall - https://github.com/timhall/ExcelHelpers
'
' Utilize OAuth2 authentication
' (Currently using client credentials flow only)
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @version 0.2.1
' @date 20121024
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

' ============================================= '
' Public Methods
' ============================================= '
Public Sub Setup(ClientId As String, ClientSecret As String, Username As String, Password As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
    Me.Username = Username
    Me.Password = Password
End Sub

Public Sub SetupTokenUrl(TokenUrl As String, Optional TokenKey As String = "access_token")
    Me.TokenUrl = TokenUrl
    Me.TokenKey = TokenKey
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub EH_IAuthenticator_BeforeExecute(request As EH_RestRequest)
    On Error GoTo ErrorHandling
    If Me.Token = "" And (Me.TokenUrl <> "" And Me.TokenKey <> "") Then
        ' Get new token
        Dim http As MSXML2.ServerXMLHTTP
        Set http = New MSXML2.ServerXMLHTTP
             
        Call http.Open("POST", CreateTokenRequest, False)
        Call http.send
        
        If http.Status <> 200 Then
            ' Error getting OAuth2 token
            Err.Raise 1, Description:="Failed to retrieve OAuth2 Token"
        End If
        
        Dim response As Dictionary
        Set response = EH_RestHelpers.ParseJSON(http.responseText)
        If Not response Is Nothing Then
            If response.Exists(Me.TokenKey) Then
                Me.Token = response(Me.TokenKey)
            End If
            
            ' (Salesforce specific, but shouldn't affect any other OAuth2 clients)
            If response.Exists("instance_url") Then
                request.BaseUrl = response("instance_url")
            End If
        Else
            Err.Raise 2, Description:="Failed to read OAuth2 Token"
        End If
    End If
    Call request.AddHeader("Authorization", CreateHeader())
    Exit Sub
    
ErrorHandling:

    If Not http Is Nothing Then Set http = Nothing
    On Error GoTo 0
End Sub

Private Sub EH_IAuthenticator_HttpOpen(http As MSXML2.IXMLHTTPRequest, request As EH_RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Open http
    Call http.Open(request.MethodName(), request.FullUrl(BaseUrl), useAsync)
End Sub

Private Function CreateHeader() As String
    CreateHeader = "OAuth " & Me.Token
End Function

Private Function CreateTokenRequest() As String
    CreateTokenRequest = Me.TokenUrl
    CreateTokenRequest = CreateTokenRequest & "&client_id=" & Me.ClientId
    CreateTokenRequest = CreateTokenRequest & "&client_secret=" & Me.ClientSecret
    CreateTokenRequest = CreateTokenRequest & "&username=" & Me.Username
    CreateTokenRequest = CreateTokenRequest & "&password=" & Me.Password
End Function

