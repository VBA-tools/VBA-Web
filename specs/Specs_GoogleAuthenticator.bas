Attribute VB_Name = "Specs_GoogleAuthenticator"
''
' Specs_GoogleAuthenticator
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Specs for GoogleAuthenticator
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

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
        .Expect(Auth.Token).ToNotBeUndefined
    End With
    
    With Specs.It("should skip login if API key is used")
        Auth.Token = ""
        Auth.APIKey = "abc"
        Auth.Login
        .Expect(Auth.Token).ToEqual ""
    End With
    
    With Specs.It("should add enabled scopes to login url")
        Auth.AddScope "http://new_scope"
        Auth.EnableScope "analytics"
        
        Dim Parts As Dictionary
        Set Parts = RestHelpers.UrlParts(Auth.LoginUrl)
        Dim Scope As String
        Scope = RestHelpers.UrlDecode(Parts("Querystring"))
        Scope = Mid$(Scope, InStr(1, Scope, "scope") + 6)
        .Expect(Scope).ToEqual "https://www.googleapis.com/auth/analytics https://www.googleapis.com/auth/userinfo.email http://new_scope"
    End With
    
    InlineRunner.RunSuite Specs
End Function
