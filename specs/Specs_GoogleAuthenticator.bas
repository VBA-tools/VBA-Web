Attribute VB_Name = "Specs_GoogleAuthenticator"
''
' Specs_GoogleAuthenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for GoogleAuthenticator
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "GoogleAuthenticator"
    
    Dim Auth As New GoogleAuthenticator
    Dim Id As String
    Dim Secret As String
    
    If Credentials.Loaded Then
        Id = Credentials.Values("Google")("id")
        Secret = Credentials.Values("Google")("secret")
    Else
        Id = InputBox("Google Client Id")
        Secret = InputBox("Google Client Secret")
    End If
    Auth.Setup Id, Secret

    With Specs.It("should login")
        Auth.Login
        .Expect(Auth.AuthorizationCode).ToNotEqual ""
    End With

    With Specs.It("should skip login if API key is used")
        Auth.Logout
        Auth.ApiKey = "abc"
        Auth.Login
        .Expect(Auth.AuthorizationCode).ToEqual ""
    End With
    
    With Specs.It("should add scopes to login url")
        Auth.AddScope "analytics"
        Auth.AddScope "http://new_scope"
        
        Dim Parts As Dictionary
        Set Parts = WebHelpers.GetUrlParts(Auth.GetLoginUrl)
        Dim Scope As String
        Scope = WebHelpers.UrlDecode(Parts("Querystring"))
        Scope = Mid$(Scope, InStr(1, Scope, "scope") + 6)
        .Expect(Scope).ToEqual "https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/analytics http://new_scope"
    End With
    
    InlineRunner.RunSuite Specs
End Function
