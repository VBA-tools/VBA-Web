VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TodoistAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Todoist Authenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for TODOist API
' https://developer.todoist.com/#oauth
' ```
'
' @class TodoistAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const AuthorizationUrl As String = "https://todoist.com/oauth/authorize"
Private Const TokenResource As String = "access_token"
Private Const BaseUrl As String = "https://todoist.com/oauth"

Private WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1
Private LoginComplete As Boolean

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public ClientSecret As String
Public Scope As String
Public RedirectURL As String
Public State As String
Public AuthorizationCode As String
Public Token As String
Public SessionCookie As New Collection

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
Public Sub Setup(ClientId As String, ClientSecret As String, RedirectURL As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
    Me.RedirectURL = RedirectURL
End Sub

''
' Login
''
Public Sub Login()
    Dim lastURL As String
    Dim newURL As String

    ' Don't need to login if we already have authorization code or token
    If Me.AuthorizationCode <> "" Or Me.Token <> "" Then
        Exit Sub
    End If

    ' Redirect users to the authorization URL
    Set IE = New InternetExplorer
    IE.Silent = True
    IE.AddressBar = False
    IE.Navigate GetLoginUrl
    IE.Visible = True

    ' The rest is handled in BeforeNavigate, but need to wait here
    Do While Not LoginComplete
        DoEvents
        newURL = IE.LocationURL
        If newURL <> lastURL Then Call Me.ParseURL(newURL)
        lastURL = newURL
    Loop

    IE.Quit
    Set IE = Nothing
    
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
    Dim Request As New WebRequest
    Request.Resource = AuthorizationUrl
    
    Request.AddQuerystringParam "client_id", Me.ClientId
    Request.AddQuerystringParam "scope", Me.Scope
    Request.AddQuerystringParam "state", Me.State
    
    GetLoginUrl = Request.FormattedResource
    Set Request = Nothing
End Function

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
Dim ck As Dictionary

    If Me.Token = "" Then
        If Me.AuthorizationCode = "" Then
            Me.Login
        End If
        
        Me.Token = Me.GetToken(Client)
    End If
    
    'add token in the beggining of the Querystring
    Request.Resource = Request.Resource & "?token=" & Me.Token
    
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
' Compares the current URL of the InternetExplorer window to the one expected
' as a return. If matches the expected one, extract the AuthorizationCode from
' it for a later exchange by a token. On sucess, sets a global boolean LoginComplete
' to True
'
' @param {String} URL
''
Public Sub ParseURL(ByVal URL As String)
    Dim UrlParts() As String

    UrlParts = Split(URL, "?")
    
    If Left(UrlParts(0), Len(Me.RedirectURL)) = Me.RedirectURL Then
        ' Parse querystring
        Dim QuerystringParams As Dictionary
        Set QuerystringParams = WebHelpers.ParseUrlEncoded(UrlParts(1))

        If QuerystringParams.Exists("error") Then
            ' TODO Handle error
        ElseIf QuerystringParams.Exists("code") Then
                If QuerystringParams("state") = Me.State Then
                    Me.AuthorizationCode = QuerystringParams("code")
                Else
                    ' TODO Handle mismatched state (unlikely but possible)
                End If
        Else
            ' TODO Handle unexpected response
        End If

        LoginComplete = True
    End If
    
End Sub


''
' Get token (for current AuthorizationCode)
'
' @internal
' @param {WebClient} Client
' @return {String}
''
Public Function GetToken(Client As WebClient) As String
    On Error GoTo Cleanup
    
    Dim TokenClient As WebClient
    Dim Request As New WebRequest
    Dim Body As New Dictionary
    Dim Response As WebResponse
    Dim Var As Variant

    
    ' Clone client (to avoid accidental interactions)
    Set TokenClient = Client.Clone
    Set TokenClient.Authenticator = Nothing
    TokenClient.BaseUrl = BaseUrl
    
    ' Prepare token request
    Request.Resource = TokenResource
    Request.Method = WebMethod.HttpPost
    Request.RequestFormat = WebFormat.FormUrlEncoded
    Request.ResponseFormat = WebFormat.Json
    
    Body.Add "code", Me.AuthorizationCode
    Body.Add "client_id", Me.ClientId
    Body.Add "client_secret", Me.ClientSecret
    Set Request.Body = Body
    
    Set Response = TokenClient.Execute(Request)
    
    If Response.StatusCode = WebStatusCode.Ok Then
        GetToken = Response.Data("access_token")
        For Each Var In Response.Cookies
            Call Me.SessionCookie.Add(Var)
        Next
    Else
        ' TODO Handle error
    End If
    
Cleanup:
    
    Set TokenClient = Nothing
    Set Request = Nothing
    Set Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        ' TODO
    End If
End Function

''
' Creates a unique state for the token login transaction
'
' @internal
''
Private Sub Class_Initialize()
    Me.State = WebHelpers.CreateNonce
End Sub
