VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FacebookAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IAuthenticator
''
' Facebook Authenticator v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for Facebook OAuth
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

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

Public Sub Setup(ApplicationId As String, ApplicationSecret As String)
    Me.ApplicationId = ApplicationId
    Me.ApplicationSecret = ApplicationSecret
End Sub

Private Sub IAuthenticator_BeforeExecute(Request As RestRequest)
    Request.AddQuerystringParam "access_token", Me.Token
End Sub

Private Sub IAuthenticator_HttpOpen( _
    http As MSXML2.IXMLHTTPRequest, _
    Request As RestRequest, _
    BaseUrl As String, _
    Optional useAsync As Boolean = False)
    
    ' Perform standard http open
    Call http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), useAsync)
End Sub

Public Sub Login()
    On Error GoTo CleanUp
    
    Dim IE As InternetExplorer
    Dim Completed As Boolean
    Dim Code As String
    Dim Token As String
    
    Set IE = New InternetExplorer
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
            Err.Raise vbObjectError + 1, "OAuthDialog", "Login failed or was denied"
        End If
    End With
    
CleanUp:
    
    If Not IE Is Nothing Then: IE.Quit
    Set IE = Nothing
    
    If Not Completed Then
        Err.Raise vbObjectError + 1, "OAuthDialog", "Login did not complete"
    ElseIf Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Err.Source, Err.Description
    ElseIf Left(Code, 5) = "Error" Then
        Err.Raise vbObjectError + 1, "OAuthDialog", Code
    Else
        ' Success!
        Me.Code = Code
        
        ' Temporarily assign token to force request
        Token = Me.Token
    End If
End Sub

Public Sub Logout()
    Me.Code = ""
    Me.Token = ""
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Function TokenRequest() As RestRequest
    Set TokenRequest = New RestRequest
    TokenRequest.Resource = "oauth/access_token"
    TokenRequest.Method = httpGET
    
    TokenRequest.AddQuerystringParam "code", Me.Code
    TokenRequest.AddQuerystringParam "client_id", Me.ApplicationId
    TokenRequest.AddQuerystringParam "client_secret", Me.ApplicationSecret
    TokenRequest.AddQuerystringParam "redirect_uri", RedirectUri
End Function

Private Function OAuthIsComplete(IE As InternetExplorer) As Boolean
    If Not IE.Busy And IE.readyState = 4 Then
        If OAuthIsSuccessful(IE) Or OAuthIsDenied(IE) Then
            OAuthIsComplete = True
        End If
    End If
End Function

Private Function OAuthIsSuccessful(IE As InternetExplorer) As Boolean
    Dim Url As String
    Url = Replace(IE.LocationURL, "https://www.facebook.com/connect/login_success.html", "")

    If (Left(Url, 5) = "#code" Or Left(Url, 6) = "#token") And Left(Url, 6) <> "?error" Then
        OAuthIsSuccessful = True
    End If
End Function

Private Function OAuthIsDenied(IE As InternetExplorer) As Boolean
    Dim Url As String
    Url = Replace(IE.LocationURL, "https://www.facebook.com/connect/login_success.html", "")

    If Left(Url, 6) = "?error" Then
        OAuthIsDenied = True
    End If
End Function
Private Function OAuthExtractCode(IE As InternetExplorer) As String
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
        OAuthExtractCode = "Error: Unrecognized token formatting"
    End If
End Function

Private Sub Class_Initialize()
    ' Email-only by default
    Me.Scope = Array("email")
    
    ' Cache token by default
    Me.CacheToken = True
End Sub

