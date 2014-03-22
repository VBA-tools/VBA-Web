VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Google Authenticator v2.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for "installed application" authentication for Google APIs
'
' - https://developers.google.com/accounts/docs/OAuth2#installed
' - https://developers.google.com/accounts/docs/OAuth2InstalledApp
'
' @implements: IAuthenticator
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
' 
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IAuthenticator
Option Explicit

Private Const AuthorizationUrl As String = "https://accounts.google.com/o/oauth2/auth"
Private Const RedirectUri As String = "urn:ietf:wg:oauth:2.0:oob"
Private pAuthorizationCode As String
Private pToken As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public ClientSecret As String
Public Scope As Variant
Public ApiKey As String

Public CacheToken As Boolean
Public LoginByDefault As Boolean

Public Property Get AuthorizationCode() As String
    If pAuthorizationCode = "" And LoginByDefault Then
        Login
    End If
    
    AuthorizationCode = pAuthorizationCode
End Property
Public Property Let AuthorizationCode(Value As String)
    pAuthorizationCode = Value
End Property

Public Property Get LoginUrl() As String
    LoginUrl = AuthorizationUrl & "?" & _
        "redirect_uri=" & RedirectUri & "&" & _
        "client_id=" & Me.ClientId & "&" & _
        "response_type=code&access_type=offline&approval_prompt=force"
    
    Dim ScopeString As String
    If IsArray(Me.Scope) Then
        ScopeString = Me.Scope(LBound(Me.Scope))
    
        Dim i As Integer
        For i = LBound(Me.Scope) + 1 To UBound(Me.Scope)
            ScopeString = ScopeString & "+" & Me.Scope(i)
        Next i
    Else
        ScopeString = Me.Scope
    End If
    LoginUrl = LoginUrl & "&scope=" & ScopeString
End Property

Public Property Get Token() As String
    On Error GoTo CleanUp
    
    If pToken = "" Or Not CacheToken Then
        Dim Client As New RestClient
        Dim Response As RestResponse
        
        ' Request a new token
        Client.BaseUrl = "https://accounts.google.com/"
        Set Response = Client.Execute(TokenRequest)
        
        ' Store token if successful, otherwise throw error
        If Response.StatusCode = Ok Then
            pToken = Response.Data("access_token")
        Else
            Err.Raise vbObjectError + Response.StatusCode, _
                Description:="Failed to load Bearer Token: " & Response.StatusCode & " - " & Response.Content
        End If
    End If
    
    ' Return token
    Token = pToken
    
CleanUp:
    
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

''
' Setup authenticator
' --------------------------------------------- '

Public Sub Setup(ClientId As String, ClientSecret As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
End Sub

Public Sub Login(Optional ApiKey As String = "")
    If ApiKey <> "" Then
        Me.ApiKey = ApiKey
    Else
        Dim Code As String
        Dim Token As String
        Dim Instructions As String
        
        Instructions = "1. Open the the login url below in your browser" & vbNewLine & _
            "2. Grant access to this application" & vbNewLine & _
            "3. Paste the authorization code below"
        
        Code = InputBox(Instructions, "Enter authorization code", Me.LoginUrl)
        
        If Code <> "" And Left(Code, 4) <> "http" Then
            Me.AuthorizationCode = Code
        Else
            Err.Raise vbObjectError + 1, _
                Description:="Invalid authorization code: " & Code
        End If
        
        ' Temporarily assign token to force request
        Token = Me.Token
    End If
End Sub

Public Sub Logout()
    Me.ApiKey = ""
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub

''
' Hook for taking action before a request is executed
'
' @param {RestClient} Client The client that is about to execute the request
' @param {RestRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IAuthenticator_BeforeExecute(ByVal Client As RestClient, ByRef Request As RestRequest)
    If Me.ApiKey <> "" Then
        Request.AddQuerystringParam "key", Me.ApiKey
    Else
        Request.AddHeader "Authorization", "Bearer " & Me.Token
    End If
End Sub

''
' Hook for taking action after request has been executed
'
' @param {RestClient} Client The client that executed request
' @param {RestRequest} Request The request that was just executed
' @param {RestResponse} Response to request
' --------------------------------------------- '

Private Sub IAuthenticator_AfterExecute(ByVal Client As RestClient, ByVal Request As RestRequest, ByRef Response As RestResponse)

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

Private Sub IAuthenticator_HttpOpen(ByRef Http As Object, ByVal Client As RestClient, ByRef Request As RestRequest, BaseUrl As String, Optional UseAsync As Boolean = False)
    
    ' Perform standard http open
    Http.Open Request.MethodName(), Request.FullUrl(BaseUrl), UseAsync
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Function TokenRequest() As RestRequest
    Set TokenRequest = New RestRequest
    TokenRequest.Resource = "o/oauth2/token"
    TokenRequest.Method = httpPOST
    
    Dim RequestBody As String
    RequestBody = "code=" & Me.AuthorizationCode & "&" & _
        "client_id=" & Me.ClientId & "&" & _
        "client_secret=" & Me.ClientSecret & "&" & _
        "redirect_uri=" & RedirectUri & "&" & _
        "grant_type=authorization_code"
    
    ' Add request body and set content-type to url-encoded
    TokenRequest.AddBodyString RequestBody
    TokenRequest.ContentType = "application/x-www-form-urlencoded"
End Function

Private Sub Class_Initialize()
    ' Cache token by default
    Me.CacheToken = True
    
    ' Only email for default scope
    Me.Scope = Array("https://www.googleapis.com/auth/userinfo.email")
End Sub



