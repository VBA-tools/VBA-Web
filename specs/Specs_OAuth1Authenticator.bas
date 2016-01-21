Attribute VB_Name = "Specs_OAuth1Authenticator"
''
' Specs_OAuth1Authenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for OAuth1Authenctiator
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "OAuth1Authenticator"
        
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Dim Auth As New OAuth1Authenticator
    Dim ConsumerKey As String
    Dim ConsumerSecret As String
    Dim Token As String
    Dim TokenSecret As String
    Dim ExpectedBaseString As String
    Dim ExpectedSignature As String
    
    If Credentials.Loaded Then
        ConsumerKey = Credentials.Values("LinkedIn")("api_key")
        ConsumerSecret = Credentials.Values("LinkedIn")("api_secret")
        Token = Credentials.Values("LinkedIn")("user_token")
        TokenSecret = Credentials.Values("LinkedIn")("user_secret")
    Else
        ConsumerKey = InputBox("Enter Consumer Key")
        ConsumerSecret = InputBox("Enter Consumer Secret")
        Token = InputBox("Enter Token")
        TokenSecret = InputBox("Enter Token Secret")
    End If
    Auth.Setup ConsumerKey, ConsumerSecret, Token, TokenSecret
    
    With Specs.It("should properly format request url")
        Client.BaseUrl = "HTTP://localhost:3000/"
        Request.Resource = "a/b/c?d=4#e"
        
        .Expect(Auth.GetRequestUrl(Client, Request)).ToEqual "http://localhost:3000/a/b/c"
    End With
    
    With Specs.It("should property format request parameters")
        Set Request = New WebRequest
        Request.Resource = "resource"
        Request.AddQuerystringParam "a", True
        Request.AddQuerystringParam "b", "abc"
        Request.AddQuerystringParam "c", 1.23
    
        .Expect(Auth.GetRequestParameters(Client, Request)).ToEqual "a=true&b=abc&c=1.23"
    End With
    
    With Specs.It("should include explicit and implicit querystring parameters")
        Client.BaseUrl = "HTTP://localhost:3000/testing"
        Set Request = New WebRequest
        Request.Resource = "?a=123&b=456"
        Request.AddQuerystringParam "c", "Howdy!"
        Request.AddQuerystringParam "d", 789
        
        .Expect(Auth.GetRequestParameters(Client, Request)).ToEqual "a=123&b=456&c=Howdy!&d=789"
    End With
    
    With Specs.It("should handle spaces in parameters correctly")
        Client.BaseUrl = "http://localhost:3000/"
        Set Request = New WebRequest
        Request.Resource = "testing"
        Request.AddQuerystringParam "a", "a b"
        
        .Expect(Auth.GetRequestParameters(Client, Request)).ToEqual "a=a%20b"
        .Expect(Client.GetFullUrl(Request)).ToEqual "http://localhost:3000/testing?a=a+b"
    End With
    
    Set Client = New WebClient
    Set Request = New WebRequest
    
    Client.BaseUrl = "HTTP://localhost:3000/"
    Request.Resource = "testing"
    Request.AddQuerystringParam "a", 123
    Request.AddQuerystringParam "b", 456
    
    Auth.Nonce = "1234"
    Auth.Timestamp = "123456789"
    
    ExpectedBaseString = "GET" & "&" & _
        WebHelpers.UrlEncode("http://localhost:3000/testing") & "&" & _
        WebHelpers.UrlEncode("a=123&b=456" & _
            "&oauth_consumer_key=" & ConsumerKey & _
            "&oauth_nonce=1234&oauth_signature_method=HMAC-SHA1&oauth_timestamp=123456789" & _
            "&oauth_token=" & Token & _
            "&oauth_version=1.0")
    
    ExpectedSignature = WebHelpers.HMACSHA1(ExpectedBaseString, ConsumerSecret & "&" & TokenSecret, "Base64")
    
    With Specs.It("should include method, resource, parameters, and oauth values in base string")
        .Expect(Auth.CreateBaseString(Auth.Nonce, Auth.Timestamp, Client, Request)).ToEqual ExpectedBaseString
    End With
    
    With Specs.It("should create signature from base and secrets with proper hashing")
        .Expect(Auth.CreateSignature(ExpectedBaseString, ConsumerSecret & "&" & TokenSecret)).ToEqual ExpectedSignature
    End With
    
    InlineRunner.RunSuite Specs
End Function

' LinkedIn Specific
' ----------------- '
Sub LinkedInSpecs()
    Dim Specs As New SpecSuite
    
    Dim Client As New WebClient
    Client.BaseUrl = "http://api.linkedin.com/v1/"
    
    Dim Auth As New OAuth1Authenticator
    Dim ConsumerKey As String
    Dim ConsumerSecret As String
    Dim Token As String
    Dim TokenSecret As String
    
    If Credentials.Loaded Then
        ConsumerKey = Credentials.Values("LinkedIn")("api_key")
        ConsumerSecret = Credentials.Values("LinkedIn")("api_secret")
        Token = Credentials.Values("LinkedIn")("user_token")
        TokenSecret = Credentials.Values("LinkedIn")("user_secret")
    Else
        ConsumerKey = InputBox("Enter Consumer Key")
        ConsumerSecret = InputBox("Enter Consumer Secret")
        Token = InputBox("Enter Token")
        TokenSecret = InputBox("Enter Token Secret")
    End If
    Auth.Setup _
        ConsumerKey:=ConsumerKey, _
        ConsumerSecret:=ConsumerSecret, _
        Token:=Token, _
        TokenSecret:=TokenSecret
    
    Set Client.Authenticator = Auth
    
    Dim Request As WebRequest
    Dim Response As WebResponse
    
    With Specs.It("should get profile")
        Set Request = New WebRequest
        Request.Resource = "people/~?format=json"
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("firstName")).ToNotBeUndefined
    End With
    
    With Specs.It("should search with space")
        Set Request = New WebRequest
        Request.Resource = "company-search?format=json"
        Request.AddQuerystringParam "keywords", "microsoft corp"
        
        Set Response = Client.Execute(Request)
        
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("companies")).ToNotBeUndefined
    End With
    
    InlineRunner.RunSuite Specs
End Sub
