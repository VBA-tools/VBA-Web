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
' Google Authenticator v2.0.5
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for "installed application" authentication for Google APIs
'
' Details:
' - https://developers.google.com/accounts/docs/OAuth2#installed
' - https://developers.google.com/accounts/docs/OAuth2InstalledApp
'
' Register for Client Id and Client Secret:
' https://console.developers.google.com/
'
' List of available scopes:
' https://developers.google.com/oauthplayground/
'
' @example
'
' Dim Auth As New GoogleAuthenticator
' Auth.Setup "Your Client Id", "Your Client Secret"
'
' ' Enable analytics scope and add YouTube scope
' Auth.EnableScope("analytics")
' Auth.AddScope("https://www.googleapis.com/auth/yt-analytics.readonly")
'
' ' Open up Google login
' Auth.Login
'
' ' or - Use your API key to access your own information
' Auth.ApiKey = "Your Api Key"
'
' ' Add authenticator to client
' Set Client.Authenticator = Auth
'
' @implements: IAuthenticator v3.*
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
Public APIKey As String
Public Scopes As Dictionary
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
    ' Use Request for Url helpers
    Dim Request As New RestRequest
    Request.Resource = AuthorizationUrl
    Request.IncludeCacheBreaker = False
    
    Request.AddQuerystringParam "redirect_uri", RedirectUri
    Request.AddQuerystringParam "client_id", Me.ClientId
    Request.AddQuerystringParam "response_type", "code"
    Request.AddQuerystringParam "access_type", "offline"
    Request.AddQuerystringParam "approval_prompt", "force"
    
    Dim ScopeString As String
    If Not Me.Scopes Is Nothing Then
        Dim Scope As Variant
        Dim UrlParts As Dictionary
        For Each Scope In Me.Scopes.Keys()
            If Me.Scopes(Scope) Then
                Set UrlParts = RestHelpers.UrlParts(CStr(Scope))
                If UrlParts("Protocol") = "" Then
                    Scope = "https://www.googleapis.com/auth/" & Scope
                End If
                
                If ScopeString = "" Then
                    ScopeString = Scope
                Else
                    ScopeString = ScopeString & " " & Scope
                End If
            End If
        Next Scope
    End If
    
    Request.AddQuerystringParam "scope", ScopeString
    
    LoginUrl = Request.FormattedResource
    RestHelpers.LogDebug LoginUrl, "GoogleAuthenticator.LoginUrl"
    Set Request = Nothing
End Property

Public Property Get Token() As String
    On Error GoTo CleanUp
    
    If pToken = "" And Me.APIKey = "" Then
        Dim Client As New RestClient
        Dim Response As RestResponse
        
        ' Request a new token
        Client.BaseUrl = "https://accounts.google.com/"
        Set Response = Client.Execute(TokenRequest)
        
        ' Store token if successful, otherwise throw error
        If Response.StatusCode = Ok Then
            Dim Data As Object
            Set Data = RestHelpers.ParseJSON(Response.Content)
            pToken = Data("access_token")
            RestHelpers.LogDebug "Received Token: " & RestHelpers.Obfuscate(pToken), "GoogleAuthenticator.Token"
        Else
            RestHelpers.LogError "Failed to load token: " & Response.StatusCode & " - " & Response.Content, "GoogleAuthenticator.Token"
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

''
' Login to Google
' --------------------------------------------- '

Public Sub Login(Optional APIKey As String = "")
    On Error GoTo CleanUp
    
    Dim IE As Object
    Dim Completed As Boolean
    Dim ErrorCode As String
    Dim Code As String
    
    Completed = True
    If APIKey <> "" Then
        Me.APIKey = APIKey
    ElseIf Me.APIKey = "" Then
        Completed = False
        Set IE = CreateObject("InternetExplorer.Application")
        With IE
            .Silent = True
            .AddressBar = False
            
            .Navigate LoginUrl
            .Visible = True
            
            Do While Not OAuthIsComplete(IE)
                ' Wait asynchronously for login to complete
                DoEvents
            Loop
            
            Completed = True
            If OAuthIsDenied(IE) Then
                ' Login failed
                RestHelpers.LogError "Login failed or was denied", "GoogleAuthenticator.Login"
                Err.Raise vbObjectError + 1, "OAuthDialog", "Login failed or was denied"
            Else
                Code = OAuthExtractCode(IE)
                If Left(Code, 5) = "Error" Then
                    RestHelpers.LogError "Login error: " & Code, "GoogleAuthenticator.Login"
                    Err.Raise vbObjectError + 1, "OAuthDialog", Code
                Else
                    ' Success!
                    Me.AuthorizationCode = Code
                    RestHelpers.LogDebug "Login success: " & Code, "GoogleAuthenticator.Login"
                    
                    ' Temporarily assign token to force request
                    Dim Token As String
                    Token = Me.Token
                End If
            End If
        End With
    End If
    
CleanUp:

    If Not IE Is Nothing Then: IE.Quit
    Set IE = Nothing
    
    If Not Completed Then
        RestHelpers.LogError "Login did not complete", "GoogleAuthenticator.Login"
        Err.Raise vbObjectError + 1, "OAuthDialog", "Login did not complete"
    ElseIf Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

''
' Logout of Google
' --------------------------------------------- '

Public Sub Logout()
    Me.APIKey = ""
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub

''
' Add scope to scopes to authorized
'
' To get a list of available scopes, visit https://developers.google.com/oauthplayground/
'
' @example
' (YouTube analytics)
' Auth.AddScope("https://www.googleapis.com/auth/yt-analytics.readonly")
'
' @param {String} Scope
' --------------------------------------------- '

Public Sub AddScope(Scope As String)
    Dim Scopes As Dictionary
    Set Scopes = Me.Scopes
    
    If Scopes.Exists(Scope) Then
        Scopes(Scope) = True
    Else
        Scopes.Add Scope, True
    End If
    
    Set Me.Scopes = Scopes
End Sub

''
' Enable scope to add to scopes to authorized
'
' Available scopes
' - calendar
' - https://www.google.com/m8/feeds/
' - analytics
' - coordinate
' - userinfo.email
' - userinfo.profile
' - https://spreadsheets.google.com/feeds/
' - urlshortener
'
' @param {String} Scope
' --------------------------------------------- '

Public Sub EnableScope(Scope As String)
    Dim Scopes As Dictionary
    Set Scopes = Me.Scopes
    
    If Scopes.Exists(Scope) Then
        Scopes(Scope) = True
    Else
        Scopes.Add Scope, True
    End If
    
    Set Me.Scopes = Scopes
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {RestClient} Client The client that is about to execute the request
' @param {RestRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IAuthenticator_BeforeExecute(ByVal Client As RestClient, ByRef Request As RestRequest)
    If Me.APIKey <> "" Then
        Request.AddQuerystringParam "key", Me.APIKey
        RestHelpers.LogDebug "Login with key: " & RestHelpers.Obfuscate(Me.APIKey), "GoogleAuthenticator.BeforeExecute"
    Else
        Request.AddHeader "Authorization", "Bearer " & Me.Token
        RestHelpers.LogDebug "Login with token: " & RestHelpers.Obfuscate(Me.Token), "GoogleAuthenticator.BeforeExecute"
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
    RestHelpers.LogDebug Request.MethodName() & Request.FullUrl(BaseUrl), "FacebookAuthenticator.HttpOpen"
End Sub

Private Function TokenRequest() As RestRequest
    Set TokenRequest = New RestRequest
    TokenRequest.Resource = "o/oauth2/token"
    TokenRequest.Method = httpPOST
    TokenRequest.Format = formurlencoded
    
    TokenRequest.AddParameter "code", Me.AuthorizationCode
    TokenRequest.AddParameter "client_id", Me.ClientId
    TokenRequest.AddParameter "client_secret", Me.ClientSecret
    TokenRequest.AddParameter "redirect_uri", RedirectUri
    TokenRequest.AddParameter "grant_type", "authorization_code"
End Function

Private Function OAuthIsComplete(IE As Object) As Boolean
    If Not IE.Busy And IE.readyState = 4 Then
        Dim UrlParts As Dictionary
        Set UrlParts = RestHelpers.UrlParts(IE.LocationURL)
        
        OAuthIsComplete = OAuthIsApproval(IE) Or OAuthIsError(IE)
    End If
End Function

Private Function OAuthIsApproval(IE As Object) As Boolean
    Dim UrlParts As Dictionary
    Set UrlParts = RestHelpers.UrlParts(IE.LocationURL)
    
    OAuthIsApproval = UrlParts("Uri") = "/o/oauth2/approval"
End Function

Private Function OAuthIsDenied(IE As Object) As Boolean
    Dim UrlParts As Dictionary
    Set UrlParts = RestHelpers.UrlParts(IE.LocationURL)
    
    If OAuthIsApproval(IE) Then
        Dim Document As Object
        Set Document = IE.Document
        
        Dim Element As Object
        For Each Element In Document.Body.all
            If Element.nodeName = "P" And Element.ID = "access_denied" Then
                OAuthIsDenied = True
                Exit Function
            End If
        Next Element
    End If
End Function

Private Function OAuthIsError(IE As Object) As Boolean
    Dim Document As Object
    Set Document = IE.Document
    If Not Document Is Nothing Then
        OAuthIsError = InStr(1, Document.Body.innerHTML, "errorCode") > 0
    End If
End Function

Private Function OAuthExtractCode(IE As Object) As String
    Dim UrlParts As Dictionary
    Set UrlParts = RestHelpers.UrlParts(IE.LocationURL)
    
    If OAuthIsApproval(IE) Then
        ' Extract authorization code
        Dim Document As Object
        Dim Element As Object
        
        Set Document = IE.Document
        For Each Element In Document.Body.all
            If Element.nodeName = "INPUT" Then
                OAuthExtractCode = Element.DefaultValue
                Exit Function
            End If
        Next Element
    Else
        ' Login error, extract error message
        OAuthExtractCode = OAuthExtractError(IE)
    End If
End Function

Private Function OAuthExtractError(IE As Object) As String
    Dim Document As Object
    Dim Element As Object
    
    Set Document = IE.Document
    For Each Element In Document.Body.all
        If Element.ID = "errorCode" Then
            OAuthExtractError = Element.innerHTML
        ElseIf Element.ID = "errorDescription" Then
            OAuthExtractError = OAuthExtractError & ", " & Element.innerHTML
            Exit Function
        End If
    Next Element
End Function

Private Sub Class_Initialize()
    ' Setup scopes
    Dim ScopeKeys As Variant
    Dim Scopes As New Dictionary
    Dim Enabled As Boolean
    Dim i As Integer
    
    ScopeKeys = Array( _
        "calendar", _
        "https://www.google.com/m8/feeds/", _
        "analytics", _
        "coordinate", _
        "userinfo.email", _
        "userinfo.profile", _
        "https://spreadsheets.google.com/feeds/", _
        "urlshortener" _
    )
    For i = LBound(ScopeKeys) To UBound(ScopeKeys)
        ' Enable only email by default
        Enabled = ScopeKeys(i) = "userinfo.email"
        Scopes.Add ScopeKeys(i), Enabled
    Next i
    Set Me.Scopes = Scopes
End Sub
