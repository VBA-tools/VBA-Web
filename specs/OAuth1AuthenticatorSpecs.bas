Attribute VB_Name = "OAuth1AuthenticatorSpecs"
''
' OAuth1AuthenticatorSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' General specs for the OAuth1Authenctiator class
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "OAuth1Authenticator"
        
    Dim Client As New RestClient
    Dim Request As New RestRequest
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
        
        .Expect(Auth.RequestUrl(Client, Request)).ToEqual "http://localhost:3000/a/b/c"
    End With
    
    With Specs.It("should property format request parameters")
        Set Request = New RestRequest
        Request.Resource = "resource"
        Request.IncludeCacheBreaker = False
        Request.AddParameter "a", True
        Request.AddParameter "b", "abc"
        Request.AddParameter "c", 1.23
    
        .Expect(Auth.RequestParameters(Client, Request)).ToEqual "a=true&b=abc&c=1.23"
    End With
    
    With Specs.It("should include explicit and implicit querystring parameters")
        Client.BaseUrl = "HTTP://localhost:3000/testing"
        Set Request = New RestRequest
        Request.IncludeCacheBreaker = False
        Request.Resource = "?a=123&b=456"
        Request.AddParameter "c", "Howdy!"
        Request.AddQuerystringParam "d", 789
        
        .Expect(Auth.RequestParameters(Client, Request)).ToEqual "a=123&b=456&c=Howdy!&d=789"
    End With
    
    Set Client = New RestClient
    Set Request = New RestRequest
    
    Client.BaseUrl = "HTTP://localhost:3000/"
    Request.Resource = "testing"
    Request.AddParameter "a", 123
    Request.AddQuerystringParam "b", 456
    Request.IncludeCacheBreaker = False
    
    Auth.Nonce = "1234"
    Auth.Timestamp = "123456789"
    
    ExpectedBaseString = "GET" & "&" & _
        UrlEncode("http://localhost:3000/testing") & "&" & _
        UrlEncode("a=123&b=456" & _
            "&oauth_consumer_key=" & ConsumerKey & _
            "&oauth_nonce=1234&oauth_signature_method=HMAC-SHA1&oauth_timestamp=123456789" & _
            "&oauth_token=" & Token & _
            "&oauth_version=1.0")
    
    ExpectedSignature = RestHelpers.Base64_HMACSHA1(ExpectedBaseString, ConsumerSecret & "&" & TokenSecret)
    
    With Specs.It("should include method, resource, parameters, and oauth values in base string")
        .Expect(Auth.CreateBaseString(Auth.Nonce, Auth.Timestamp, Client, Request)).ToEqual ExpectedBaseString
    End With
    
    With Specs.It("should create signature from base and secrets with proper hashing")
        .Expect(Auth.CreateSignature(ExpectedBaseString, Auth.CreateSigningKey())).ToEqual ExpectedSignature
    End With
    
    InlineRunner.RunSuite Specs
End Function
