Attribute VB_Name = "GoogleAuthenticatorSpecs"
''
' GoogleAuthenticatorSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' General specs for the GoogleAuthenticator class
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "DigestAuthenticator"
    
    Dim Auth As New GoogleAuthenticator
    Dim ID As String
    Dim Secret As String
    
    If Credentials.Loaded Then
        ID = Credentials.Values("Google")("id")
        Secret = Credentials.Values("Google")("secret")
    Else
        ID = InputBox("Google Client Id")
        Secret = InputBox("Google Client Secret")
    End If
    Auth.Setup ID, Secret
    
    With Specs.It("should login")
        Auth.Login
        .Expect(Auth.Token).ToBeDefined
    End With
    
    With Specs.It("should skip login if API key is used")
        Auth.Token = ""
        Auth.ApiKey = "abc"
        Auth.Login
        .Expect(Auth.Token).ToEqual ""
    End With
    
    With Specs.It("should add enabled scopes to login url")
        Auth.AddScope "http://new_scope"
        Auth.EnableScope "analytics"
        
        Dim Parts As Dictionary
        Set Parts = RestHelpers.UrlParts(Auth.LoginUrl)
        Dim Scope As String
        Scope = RestHelpers.URLDecode(Parts("Querystring"))
        Scope = Mid$(Scope, InStr(1, Scope, "scope") + 6)
        .Expect(Scope).ToEqual "https://www.googleapis.com/auth/analytics+https://www.googleapis.com/auth/userinfo.email+http://new_scope"
    End With
    
    InlineRunner.RunSuite Specs
End Function




