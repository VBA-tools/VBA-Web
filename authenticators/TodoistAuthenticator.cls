VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TodoistAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Todoist Authenticator v3.0.8
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'     Mauricio Souza (mauriciojxs@yahoo.com.br)
'
' Custom IWebAuthenticator for OAuth2 redirect authentication for Todoist API
'
' Details:
' - https://developer.todoist.com/#oauth
'
' Note:
' For the redirect url, any valid url seems to work (as only the querystring is used in authentication)
' e.g. www.example.com, www.yourcompany.com, etc.
'
' The redirect url must match the value set in the "App Management Console"
'
' Developers:
' - ClientId, ClientSecret, RedirectUrl: https://todoist.com/app_console/
' - RedirectUrl: any valid url seems to work (only the querystring is used in authentication)
'   e.g. www.example.com, www.yourcompany.com, etc.
' - List of available scopes: https://developer.todoist.com/#oauth
'
' Errors:
' 11040 / 80042b20 / -2147210464 - Error logging in
' 11041 / 80042b21 / -2147210463 - Error retrieving token
' 11043 / 80042b23 / -2147210461 - State response from API does not match state provided to API, the request may have been compromised
'
' @example
' ```VB.net
' Dim Auth As New TodoistAuthenticator
' Auth.Setup "Your ClientId", "Your ClientSecret", "Your RedirectURL"
'
' ' Set API scope (comma-separated)
' Auth.Scope = "data:read,data:delete"
'
' Set Client.Authenticator = Auth
' ```
'
' @class TodoistAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com, mauriciojxs@yahoo.com.br
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_AuthorizationUrl As String = "https://todoist.com/oauth/authorize"
Private Const auth_TokenResource As String = "access_token"
Private Const auth_BaseUrl As String = "https://todoist.com/oauth"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public ClientSecret As String
Public RedirectUrl As String
Public Scope As String
Public State As String
Public AuthorizationCode As String
Public Token As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} ClientId
' @param {String} ClientSecret
' @param {String} RedirectURL
''
Public Sub Setup(ClientId As String, ClientSecret As String, RedirectUrl As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
    Me.RedirectUrl = RedirectUrl
End Sub

''
' Login
''
Public Sub Login()
    On Error GoTo auth_ErrorHandling

    ' Don't need to login if we already have authorization code or token
    If Me.AuthorizationCode <> "" Or Me.Token <> "" Then
        Exit Sub
    End If

    Dim auth_Completed As Boolean
    If Me.State = "" Then
        Me.State = WebHelpers.CreateNonce
    End If

#If Mac Then
    ' User retrieves and pastes token for Mac login
    Dim auth_Response As String
    auth_Completed = True

    auth_Response = VBA.InputBox("To login to Todoist, insert your API token from" & vbNewLine & _
        "todoist.com > Settings > Todoist Settings > Account > API token", _
        Title:="Insert Todoist API Token")

    If auth_Response = "" Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login was cancelled"
    End If

    ' Success!
    Me.Token = auth_Response
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
        WebHelpers.LogDebug "Login succeeded: " & Me.AuthorizationCode, "TodoistAuthenticator.Login"
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

    WebHelpers.LogError auth_ErrorDescription, "TodoistAuthenticator.Login", 11040 + vbObjectError
    Err.Raise 11040 + vbObjectError, "TodoistAuthenticator.Login", auth_ErrorDescription
End Sub

''
' Logout
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub

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

    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "scope", Me.Scope
    auth_Request.AddQuerystringParam "state", Me.State

    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    If Me.Token = "" Then
        If Me.AuthorizationCode = "" Then
            Me.Login
        End If

        Me.Token = Me.GetToken(Client)
    End If

    ' Add token as querystring param
    Request.AddQuerystringParam "token", Me.Token
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
    Dim auth_Cookie As Variant

    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = auth_BaseUrl

    ' Prepare token request
    auth_Request.Resource = auth_TokenResource
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json

    auth_Body.Add "code", Me.AuthorizationCode
    auth_Body.Add "client_id", Me.ClientId
    auth_Body.Add "client_secret", Me.ClientSecret
    Set auth_Request.Body = auth_Body

    Set auth_Response = auth_TokenClient.Execute(auth_Request)

    If auth_Response.StatusCode = WebStatusCode.Ok Then
        GetToken = auth_Response.Data(auth_TokenResource)
    Else
        Err.Raise 11041 + vbObjectError, "TodoistAuthenticator.GetToken", _
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

        WebHelpers.LogError auth_ErrorDescription, "TodoistAuthenticator.GetToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "TodoistAuthenticator.GetToken", auth_ErrorDescription
    End If
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function auth_LoginIsComplete(auth_IE As Object) As Boolean
    If Not auth_IE.Busy And auth_IE.ReadyState = 4 Then
        auth_LoginIsComplete = VBA.InStr(1, auth_IE.LocationURL, Me.RedirectUrl)
    End If
End Function

Private Function auth_LoginIsApproval(auth_IE As Object) As Boolean
    Dim auth_UrlParts As Dictionary
    Dim auth_Querystring As Dictionary
    Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
    Set auth_Querystring = WebHelpers.ParseUrlEncoded(auth_UrlParts("Querystring"))

    auth_LoginIsApproval = Not auth_Querystring.Exists("error")
End Function

Private Function auth_LoginIsDenied(auth_IE As Object) As Boolean
    Dim auth_UrlParts As Dictionary
    Dim auth_Querystring As Dictionary
    Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
    Set auth_Querystring = WebHelpers.ParseUrlEncoded(auth_UrlParts("Querystring"))

    If auth_Querystring.Exists("error") Then
        If auth_Querystring("error") = "access_denied" Then
            auth_LoginIsDenied = True
        End If
    End If
End Function

Private Function auth_LoginIsError(auth_IE As Object) As Boolean
    Dim auth_UrlParts As Dictionary
    Dim auth_Querystring As Dictionary
    Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
    Set auth_Querystring = WebHelpers.ParseUrlEncoded(auth_UrlParts("Querystring"))

    If auth_Querystring.Exists("error") Then
        If auth_Querystring("error") <> "access_denied" Then
            auth_LoginIsError = True
        End If
    End If
End Function

Private Function auth_LoginExtractCode(auth_IE As Object) As String
    Dim auth_UrlParts As Dictionary
    Dim auth_Querystring As Dictionary
    Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
    Set auth_Querystring = WebHelpers.ParseUrlEncoded(auth_UrlParts("Querystring"))

    If auth_Querystring("state") <> Me.State Then
        Err.Raise 11043 + vbObjectError, "TodoistAuthenticator.LoginExtractCode", "State response from API does not match state provided to API, the request may have been compromised"
    Else
        auth_LoginExtractCode = auth_Querystring("code")
    End If
End Function

Private Function auth_LoginExtractError(auth_IE As Object) As String
    Dim auth_UrlParts As Dictionary
    Dim auth_Querystring As Dictionary
    Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
    Set auth_Querystring = WebHelpers.ParseUrlEncoded(auth_UrlParts("Querystring"))

    auth_LoginExtractError = auth_Querystring("error")
End Function
