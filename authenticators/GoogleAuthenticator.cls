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
' Google Authenticator v3.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for "installed application" authentication for Google APIs
'
' Details:
' - https://developers.google.com/accounts/docs/OAuth2#installed
' - https://developers.google.com/accounts/docs/OAuth2InstalledApp
'
' Developers:
' - Register for Client Id and Client Secret: https://console.developers.google.com/
' - List of available scopes: https://developers.google.com/oauthplayground/
'
' Errors:
' 11040 / 80042b20 / -2147210464 - Error logging in
' 11041 / 80042b21 / -2147210463 - Error retrieving token
'
' @example
' ```VB.net
' Dim Auth As New GoogleAuthenticator
' Auth.Setup "Your Client Id", "Your Client Secret"
'
' ' Add Google Analytics and Gmail scopes
' ' (https://www.googleapis.com/auth/ is added automatically when no domain is specified)
' Auth.AddScope "analytics" ' -> https://www.googleapis.com/auth/analytics
' Auth.AddScope "https://mail.google.com/"
'
' ' Manually open up Google login
' ' (called automatically on first request otherwise)
' Auth.Login
'
' ' alternatively, use your API key to bypass login process
' Auth.ApiKey = "Your API Key"
'
' ' Add authenticator to client
' Set Client.Authenticator = Auth
' ```
'
' @class GoogleAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_AuthorizationUrl As String = "https://accounts.google.com/o/oauth2/auth"
Private Const auth_RedirectUrl As String = "urn:ietf:wg:oauth:2.0:oob"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public ClientSecret As String
Public ApiKey As String
Public AuthorizationCode As String
Public Token As String
Public Scopes As Variant

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} ClientId
' @param {String} ClientSecret
''
Public Sub Setup(ClientId As String, ClientSecret As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
End Sub

''
' Login to Google
''
Public Sub Login()
    On Error GoTo auth_ErrorHandling
    
    ' No need to login if API key, authorization code, or token have been set
    If Me.ApiKey <> "" Or Me.AuthorizationCode <> "" Or Me.Token <> "" Then
        Exit Sub
    End If
    
    Dim auth_Completed As Boolean
    auth_Completed = True
    
#If Mac Then
    
    ' Mac login opens dialog and then user copy-paste's authorization code into InputBox
    Dim auth_Result As ShellResult
    Dim auth_Response As String
    
    auth_Result = WebHelpers.ExecuteInShell("open " & WebHelpers.PrepareTextForShell(Me.GetLoginUrl))
    
    If auth_Result.ExitCode <> 0 Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Unable to open browser"
    End If
    
    auth_Response = VBA.InputBox("Opening Google Login..." & vbNewLine & vbNewLine & _
        "After you've logged in, copy the code from the browser and paste it here to authorize this application", _
        Title:="Logging in...")
    
    If auth_Response = "" Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login was cancelled"
    End If
    
    ' Success!
    Me.AuthorizationCode = auth_Response
    
#Else

    ' Windows login uses IE to automate retrieving authorization code for user
    On Error GoTo auth_Cleanup
    
    Dim auth_IE As Object
    auth_Completed = False
    
    Set auth_IE = CreateObject("InternetExplorer.Application")
    auth_IE.Silent = True
    auth_IE.AddressBar = False
    auth_IE.Navigate Me.GetLoginUrl
    auth_IE.Visible = True
        
    ' Wait for login to complete
    Do While Not auth_LoginIsComplete(auth_IE)
        DoEvents
    Loop
    auth_Completed = True
    
    If auth_LoginIsDenied(auth_IE) Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login failed or was denied"
    ElseIf auth_LoginIsError(auth_IE) Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login error: " & auth_LoginExtractError(auth_IE)
    End If
    
    ' Success!
    Me.AuthorizationCode = auth_LoginExtractCode(auth_IE)
    
auth_Cleanup:

    If Not auth_IE Is Nothing Then: auth_IE.Quit
    Set auth_IE = Nothing
    
#End If

    If Err.Number = 0 And auth_Completed Then
        WebHelpers.LogDebug "Login succeeded: " & Me.AuthorizationCode, "GoogleAuthenticator.Login"
        Exit Sub
    End If

auth_ErrorHandling:
    
    Dim auth_ErrorDescription As String
    
    auth_ErrorDescription = "An error occurred while logging in." & vbNewLine
    If Err.Number <> 0 Then
        If Err.Number - vbObjectError <> 11040 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
    Else
        auth_ErrorDescription = auth_ErrorDescription & "Login did not complete"
    End If
    auth_ErrorDescription = auth_ErrorDescription & Err.Description

    WebHelpers.LogError auth_ErrorDescription, "GoogleAuthenticator.Login", 11040 + vbObjectError
    Err.Raise 11040 + vbObjectError, "GoogleAuthenticator.Login", auth_ErrorDescription
End Sub

''
' Logout
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub

''
' Add scope to authorized scopes
'
' - To get a list of available scopes, visit https://developers.google.com/oauthplayground/
' - As a shortcut, if a domain isn't given, https://www.googleapis.com/auth/ is automatically added
'
' @example
' Auth.AddScope "yt-analytics.readonly" ' -> https://www.googleapis.com/auth/yt-analytics.readonly
' Auth.AddScope "https://mail.google.com/"
'
' @param {String} Scope
''
Public Sub AddScope(Scope As String)
    Dim auth_Scopes As Variant
    
    ' Prepare scopes array
    auth_Scopes = Me.Scopes
    If VBA.IsEmpty(auth_Scopes) Then
        ReDim auth_Scopes(0 To 0)
    Else
        ReDim Preserve auth_Scopes(0 To UBound(auth_Scopes) + 1)
    End If
    
    ' Add standard domain if it hasn't been set
    If VBA.Left$(Scope, 4) <> "http" And Not VBA.InStr(1, Scope, "://") Then
        Scope = "https://www.googleapis.com/auth/" & Scope
    End If
    
    auth_Scopes(UBound(auth_Scopes)) = Scope
    Me.Scopes = auth_Scopes
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    If Me.ApiKey <> "" Then
        Request.AddQuerystringParam "key", Me.ApiKey
    Else
        If Me.Token = "" Then
            If Me.AuthorizationCode = "" Then
                Me.Login
            End If
            
            Me.Token = Me.GetToken(Client)
        End If
    
        Request.SetHeader "Authorization", "Bearer " & Me.Token
    End If
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

''
' Get token (for current AuthorizationCode)
'
' @internal
' @param {WebClient} Client
' @return {String}
''
Public Function GetToken(Client As WebClient) As String
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://accounts.google.com/"
    
    ' Prepare token request
    auth_Request.Resource = "o/oauth2/token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    
    auth_Body.Add "code", Me.AuthorizationCode
    auth_Body.Add "client_id", Me.ClientId
    auth_Body.Add "client_secret", Me.ClientSecret
    auth_Body.Add "redirect_uri", auth_RedirectUrl
    auth_Body.Add "grant_type", "authorization_code"
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        GetToken = auth_Response.Data("access_token")
    Else
        Err.Raise 11041 + vbObjectError, "GoogleAuthenticator.GetToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "GoogleAuthenticator.GetToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "GoogleAuthenticator.GetToken", auth_ErrorDescription
    End If
End Function

''
' Get login url for current scopes
'
' @internal
' @return {String}
''
Public Function GetLoginUrl() As String
    ' Use Request for Url helpers
    Dim auth_Request As New WebRequest
    auth_Request.Resource = auth_AuthorizationUrl
    
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "response_type", "code"
    auth_Request.AddQuerystringParam "access_type", "offline"
    auth_Request.AddQuerystringParam "approval_prompt", "force"
    
    If Not VBA.IsEmpty(Me.Scopes) Then
        auth_Request.AddQuerystringParam "scope", VBA.Join(Me.Scopes, " ")
    Else
        auth_Request.AddQuerystringParam "scope", ""
    End If
    
    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function auth_LoginIsComplete(auth_IE As Object) As Boolean
    If Not auth_IE.Busy And auth_IE.ReadyState = 4 Then
        auth_LoginIsComplete = auth_LoginIsApproval(auth_IE) Or auth_LoginIsError(auth_IE)
    End If
End Function

Private Function auth_LoginIsApproval(auth_IE As Object) As Boolean
    Dim auth_UrlParts As Dictionary
    Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
    
    auth_LoginIsApproval = auth_UrlParts("Path") = "/o/oauth2/approval"
End Function

Private Function auth_LoginIsDenied(auth_IE As Object) As Boolean
    Dim auth_Document As Object
    Dim auth_Element As Object

    If auth_LoginIsApproval(auth_IE) Then
        For Each auth_Element In auth_IE.Document.Body.All
            If VBA.UCase(auth_Element.NodeName) = "P" And auth_Element.Id = "access_denied" Then
                auth_LoginIsDenied = True
                Exit Function
            End If
        Next auth_Element
    End If
End Function

Private Function auth_LoginIsError(auth_IE As Object) As Boolean
    auth_LoginIsError = InStr(1, auth_IE.Document.Body.innerHTML, "errorCode") > 0
End Function

Private Function auth_LoginExtractCode(auth_IE As Object) As String
    Dim auth_Element As Object
    
    If auth_LoginIsApproval(auth_IE) Then
        ' Extract authorization code
        For Each auth_Element In auth_IE.Document.Body.All
            If VBA.UCase(auth_Element.NodeName) = "INPUT" Then
                auth_LoginExtractCode = auth_Element.DefaultValue
                Exit Function
            End If
        Next auth_Element
    End If
End Function

Private Function auth_LoginExtractError(auth_IE As Object) As String
    Dim auth_Element As Object
    
    For Each auth_Element In auth_IE.Document.Body.All
        If auth_Element.Id = "errorCode" Then
            auth_LoginExtractError = auth_Element.innerHTML
        ElseIf auth_Element.Id = "errorDescription" Then
            auth_LoginExtractError = auth_LoginExtractError & ", " & auth_Element.innerHTML
            Exit Function
        End If
    Next auth_Element
End Function

Private Sub Class_Initialize()
    Me.Scopes = Array("https://www.googleapis.com/auth/userinfo.email")
End Sub
