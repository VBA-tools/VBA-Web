VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FacebookAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Facebook Authenticator v3.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for Facebook
' Note: Windows-only currently
'
' Developers:
' - Scopes: https://developers.facebook.com/docs/facebook-login/permissions/v2.2
'
' Errors:
' 11040 / 80042b20 / -2147210464 - Error logging in
' 11041 / 80042b21 / -2147210463 - Error retreiving token
' 11042 / 80042b22 / -2147210462 - Error parsing authorization code
'
' @example
' ```VB.net
' Dim Auth As New FacebookAuthenticator
' Auth.Setup "Your Application Id", "Your Application Secret"
'
' ' Add user location scope
' Auth.AddScope "user_location"
'
' ' Manually open Facebook login
' ' (called automatically on first request otherwise)
' Auth.Login
'
' ' Add authenticator to client
' Set Client.Authenticator = Auth
' ```
'
' @class FacebookAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_AuthorizationUrl As String = "https://www.facebook.com/dialog/oauth"
Private Const auth_RedirectUrl As String = "https://www.facebook.com/connect/login_success.html"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ApplicationId As String
Public ApplicationSecret As String
Public AuthorizationCode As String
Public Token As String
Public Scopes As Variant

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} ApplicationId
' @param {String} ApplicationSecret
''
Public Sub Setup(ApplicationId As String, ApplicationSecret As String)
    Me.ApplicationId = ApplicationId
    Me.ApplicationSecret = ApplicationSecret
End Sub

''
' Login to Facebook
''
Public Sub Login()
    On Error GoTo auth_ErrorHandling
    
    Dim auth_Completed As Boolean

#If Mac Then
    
    ' Mac login
    Err.Raise 11040 + vbObjectError, Description:="Facebook authentication is not currently supported on Mac"
    
#Else
    
    ' Windows login uses IE to automate retrieving authorization code for user
    On Error GoTo auth_Cleanup
    
    Dim auth_IE As Object
    
    Set auth_IE = CreateObject("InternetExplorer.Application")
    auth_IE.Silent = True
    auth_IE.AddressBar = False
    auth_IE.Visible = True
    auth_IE.Navigate Me.GetLoginUrl
        
    ' Wait for login to complete
    Do While Not auth_LoginIsComplete(auth_IE)
        DoEvents
    Loop
    auth_Completed = True
    
    If auth_LoginIsDenied(auth_IE) Then
        Err.Raise 11040 + vbObjectError, Description:="Login failed or was denied"
    ElseIf auth_LoginIsError(auth_IE) Then
        Err.Raise 11040 + vbObjectError, Description:="Login error: " & auth_LoginExtractError(auth_IE)
    End If
    
    ' Success!
    Me.AuthorizationCode = auth_LoginExtractCode(auth_IE)
    
auth_Cleanup:
    
    If Not auth_IE Is Nothing Then: auth_IE.Quit
    Set auth_IE = Nothing
    
#End If
    
    If Err.Number = 0 And auth_Completed Then
        WebHelpers.LogDebug "Login succeeded: " & Me.AuthorizationCode, "FacebookAuthenticator.Login"
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
            
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    Else
        auth_ErrorDescription = auth_ErrorDescription & "Login did not complete"
    End If

    WebHelpers.LogError auth_ErrorDescription, "FacebookAuthenticator.Login", 11040 + vbObjectError
    Err.Raise 11040 + vbObjectError, "FacebookAuthenticator.Login", auth_ErrorDescription
End Sub

''
' Logout of Facebook
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.Token = ""
End Sub

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

    Request.AddQuerystringParam "access_token", Me.Token
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
' Add scope to authorized scopes
'
' - To get a list of available scopes, visit https://developers.facebook.com/docs/facebook-login/permissions/v2.2
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
    
    auth_Scopes(UBound(auth_Scopes)) = Scope
    Me.Scopes = auth_Scopes
End Sub

''
' Get token with given Client
'
' @internal
' @param {WebClient} Client
' @return {String}
''
Public Function GetToken(Client As WebClient) As String
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://graph.facebook.com/"
    
    ' Prepare token request
    auth_Request.Resource = "oauth/access_token"
    auth_Request.Method = WebMethod.HttpGet
    auth_Request.ResponseFormat = WebFormat.FormUrlEncoded
    
    auth_Request.AddQuerystringParam "code", Me.AuthorizationCode
    auth_Request.AddQuerystringParam "client_id", Me.ApplicationId
    auth_Request.AddQuerystringParam "client_secret", Me.ApplicationSecret
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    
    ' Request a new token
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    ' Store token if successful, otherwise throw error
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        GetToken = auth_Response.Data("access_token")
        WebHelpers.LogDebug "Received token: " & GetToken, "FacebookAuthenticator.GetToken"
    Else
        Err.Raise 11041 + vbObjectError, Description:=auth_Response.StatusCode & ": " & auth_Response.Content
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
    
        WebHelpers.LogError auth_ErrorDescription, "FacebookAuthenticator.GetToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "FacebookAuthenticator.GetToken", auth_ErrorDescription
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
    
    auth_Request.AddQuerystringParam "client_id", Me.ApplicationId
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    auth_Request.AddQuerystringParam "response_type", "code token"
    auth_Request.AddQuerystringParam "scope", VBA.Join(Me.Scopes, ",")
    
    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function auth_LoginIsComplete(auth_IE As Object) As Boolean
    If Not auth_IE.Busy And auth_IE.ReadyState = 4 Then
        If auth_LoginIsApproval(auth_IE) Or auth_LoginIsDenied(auth_IE) Or auth_LoginIsError(auth_IE) Then
            auth_LoginIsComplete = True
        End If
    End If
End Function

Private Function auth_LoginIsApproval(auth_IE As Object) As Boolean
    Dim auth_UrlParts As Dictionary

    If VBA.InStr(1, auth_IE.LocationURL, auth_RedirectUrl) > 0 Then
        Set auth_UrlParts = WebHelpers.GetUrlParts(auth_IE.LocationURL)
        
        If VBA.InStr(1, auth_UrlParts("Hash"), "code") > 0 Or VBA.InStr(1, auth_UrlParts("Hash"), "token") > 0 Then
            auth_LoginIsApproval = True
        End If
    End If
End Function

Private Function auth_LoginIsDenied(auth_IE As Object) As Boolean
    If VBA.InStr(1, auth_IE.LocationURL, auth_RedirectUrl) > 0 Then
        If VBA.InStr(1, WebHelpers.GetUrlParts(auth_IE.LocationURL)("Querystring"), "error=access_denied") > 0 Then
            auth_LoginIsDenied = True
        End If
    End If
End Function

Private Function auth_LoginIsError(auth_IE As Object) As Boolean
    If VBA.InStr(1, auth_IE.LocationURL, auth_RedirectUrl) > 0 Then
        If VBA.InStr(1, WebHelpers.GetUrlParts(auth_IE.LocationURL)("Querystring"), "error") > 0 Then
            auth_LoginIsError = True
        End If
    ElseIf VBA.InStr(1, auth_IE.LocationURL, auth_AuthorizationUrl) > 0 Then
        ' There are minimal distinguishing marks for errors,
        ' Look for Warning without Cancel Okay of confirmation prompt
        ' TODO Find less brittle approach
        If VBA.InStr(1, auth_IE.Document.Body.innerText, "Warning") > 0 And Not VBA.InStr(1, auth_IE.Document.Body.innerText, "CancelOkay") > 0 Then
            auth_LoginIsError = True
        End If
    End If
End Function

Private Function auth_LoginExtractCode(auth_IE As Object) As String
    On Error GoTo auth_ErrorHandling
    
    Dim auth_Data As Dictionary
    
    Set auth_Data = WebHelpers.ParseUrlEncoded(WebHelpers.GetUrlParts(auth_IE.LocationURL)("Hash"))
    auth_LoginExtractCode = auth_Data("code")
    
    If auth_LoginExtractCode = "" Then
        Err.Raise 11042 + vbObjectError, Description:="Authorization code not found in " & auth_IE.LocationURL
    End If
    
    Exit Function
    
auth_ErrorHandling:

    Dim auth_ErrorDescription As String
        
    auth_ErrorDescription = "An error occurred while parsing authorization code." & vbNewLine
    If Err.Number - vbObjectError <> 11042 Then
        auth_ErrorDescription = auth_ErrorDescription & _
            Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
    End If
    auth_ErrorDescription = auth_ErrorDescription & Err.Description

    WebHelpers.LogError auth_ErrorDescription, "FacebookAuthenticator.LoginExtractCode", 11042 + vbObjectError
    Err.Raise 11042 + vbObjectError, "FacebookAuthenticator.LoginExtractCode", auth_ErrorDescription
End Function

Private Function auth_LoginExtractError(auth_IE As Object) As String
    On Error GoTo auth_ErrorHandling
    
    If VBA.InStr(1, auth_IE.LocationURL, auth_RedirectUrl) > 0 Then
        Dim auth_QuerystringParts() As String
        
        auth_QuerystringParts = VBA.Split(WebHelpers.GetUrlParts(auth_IE.LocationURL)("Querystring"), "&")
        auth_LoginExtractError = VBA.Replace(VBA.Join(auth_QuerystringParts, vbNewLine), "=", ": ")
    ElseIf VBA.InStr(1, auth_IE.LocationURL, auth_AuthorizationUrl) > 0 Then
        auth_LoginExtractError = auth_IE.Document.Body.FirstChild.innerText
    End If
    
    Exit Function
    
auth_ErrorHandling:

    auth_LoginExtractError = "An unknown error occurred"
End Function

Private Sub Class_Initialize()
    ' Email-only by default
    Me.Scopes = Array("email")
End Sub

