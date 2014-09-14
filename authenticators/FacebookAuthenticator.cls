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
' Facebook Authenticator v2.0.6
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for Facebook OAuth
'
' @implements: IAuthenticator v3.*
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IAuthenticator
Option Explicit

Private Const RedirectUri As String = "https://www.facebook.com/connect/login_success.html"
Private pToken As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ApplicationId As String
Public ApplicationSecret As String

Public CacheToken As Boolean
Public Code As String
Public Scope As Variant

Public Property Get Token() As String
    On Error GoTo CleanUp

    If (pToken = "" Or Not CacheToken) And Code <> "" Then
        Dim Client As New RestClient
        Dim Response As RestResponse
        
        ' Request a new token
        Client.BaseUrl = "https://graph.facebook.com"
        Set Response = Client.Execute(TokenRequest())
        
        ' Store token if successful, otherwise throw error
        If Response.StatusCode = Ok Then
            Dim StartIndex As Integer
            Dim EndIndex As Integer
            
            StartIndex = 14
            EndIndex = InStr(StartIndex, Response.Content, "&expires=")
            
            pToken = Mid$(Response.Content, StartIndex, EndIndex - StartIndex)
            RestHelpers.LogDebug "Received token: " & RestHelpers.Obfuscate(pToken), "FacebookAuthenticator.Token"
        Else
            RestHelpers.LogError "Failed to load token: " & Response.StatusCode & " - " & Response.Content, "FacebookAuthenticator.Token"
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

Public Property Get LoginUrl()
    LoginUrl = "https://www.facebook.com/dialog/oauth?" & _
        "client_id=" & Me.ApplicationId & _
        "&redirect_uri=" & RedirectUri & _
        "&response_type=" & "code token"
        
    Dim ScopeString As String
    If IsArray(Me.Scope) Then
        ScopeString = Me.Scope(LBound(Me.Scope))
    
        Dim i As Integer
        For i = LBound(Me.Scope) + 1 To UBound(Me.Scope)
            ScopeString = ScopeString & "," & Me.Scope(i)
        Next i
    Else
        ScopeString = Me.Scope
    End If
    LoginUrl = LoginUrl & "&scope=" & ScopeString
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup authenticator
' --------------------------------------------- '

Public Sub Setup(ApplicationId As String, ApplicationSecret As String)
    Me.ApplicationId = ApplicationId
    Me.ApplicationSecret = ApplicationSecret
End Sub

''
' Login to Facebook
' --------------------------------------------- '

Public Sub Login()
    On Error GoTo CleanUp
    
    Dim IE As Object
    Dim Completed As Boolean
    Dim Code As String
    Dim Token As String
    
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
        If OAuthIsSuccessful(IE) Then
            Code = OAuthExtractCode(IE)
        Else
            ' Login failed
            RestHelpers.LogError "Facebook login failed or was denied", "FacebookAuthenticator.Login"
            Err.Raise vbObjectError + 1, "OAuthDialog", "Login failed or was denied"
        End If
    End With
    
CleanUp:
    
    If Not IE Is Nothing Then: IE.Quit
    Set IE = Nothing
    
    If Not Completed Then
        RestHelpers.LogError "Facebook login did not complete", "FacebookAuthenticator.Login"
        Err.Raise vbObjectError + 1, "OAuthDialog", "Login did not complete"
    ElseIf Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Err.Source, Err.Description
    ElseIf Left(Code, 5) = "Error" Then
        RestHelpers.LogError "Facebook login returned error: " & Code, "FacebookAuthenticator.Login"
        Err.Raise vbObjectError + 1, "OAuthDialog", Code
    Else
        RestHelpers.LogDebug "Successfully logged in: " & Code, "FacebookAuthenticator.Login"
        ' Success!
        Me.Code = Code
        
        ' Temporarily assign token to force request
        Token = Me.Token
    End If
End Sub

''
' Logout of Facebook
' --------------------------------------------- '

Public Sub Logout()
    Me.Code = ""
    Me.Token = ""
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
    Request.AddQuerystringParam "access_token", Me.Token
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

Private Function TokenRequest() As RestRequest
    Set TokenRequest = New RestRequest
    TokenRequest.Resource = "oauth/access_token"
    TokenRequest.Method = httpGET
    
    TokenRequest.AddQuerystringParam "code", Me.Code
    TokenRequest.AddQuerystringParam "client_id", Me.ApplicationId
    TokenRequest.AddQuerystringParam "client_secret", Me.ApplicationSecret
    TokenRequest.AddQuerystringParam "redirect_uri", RedirectUri
End Function

Private Function OAuthIsComplete(IE As Object) As Boolean
    If Not IE.Busy And IE.readyState = 4 Then
        If OAuthIsSuccessful(IE) Or OAuthIsDenied(IE) Then
            OAuthIsComplete = True
        End If
    End If
End Function

Private Function OAuthIsSuccessful(IE As Object) As Boolean
    Dim Url As String
    Url = Replace(IE.LocationURL, "https://www.facebook.com/connect/login_success.html", "")

    If (Left(Url, 5) = "#code" Or Left(Url, 6) = "#token") And Left(Url, 6) <> "?error" Then
        OAuthIsSuccessful = True
    End If
End Function

Private Function OAuthIsDenied(IE As Object) As Boolean
    Dim Url As String
    Url = Replace(IE.LocationURL, "https://www.facebook.com/connect/login_success.html", "")

    If Left(Url, 6) = "?error" Then
        OAuthIsDenied = True
    End If
End Function
Private Function OAuthExtractCode(IE As Object) As String
    Dim Url As String
    Dim StartIndex As Integer
    Dim EndIndex As Integer

    Url = Replace(IE.LocationURL, "https://www.facebook.com/connect/login_success.html", "")
    StartIndex = InStr(1, Url, "code", vbTextCompare)
    StartIndex = StartIndex + 5

    EndIndex = InStr(StartIndex, Url, "access_token", vbTextCompare)
    EndIndex = EndIndex - 1

    If StartIndex >= 0 And EndIndex > StartIndex Then
        OAuthExtractCode = Mid$(Url, StartIndex, EndIndex - StartIndex)
    Else
        RestHelpers.LogError "Unrecognized token format: " & Url, "FacebookAuthenticator.OAuthExtractCode"
        OAuthExtractCode = "Error: Unrecognized token formatting"
    End If
End Function

Private Sub Class_Initialize()
    ' Email-only by default
    Me.Scope = Array("email")
    
    ' Cache token by default
    Me.CacheToken = True
End Sub

