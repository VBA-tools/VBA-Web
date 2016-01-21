Attribute VB_Name = "Specs_DigestAuthenticator"
''
' Specs_DigestAuthenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for DigestAuthenticator
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "DigestAuthenticator"
    
    Dim web_CrLf As String
    web_CrLf = VBA.Chr$(13) & VBA.Chr$(10)
    
    Dim Auth As New DigestAuthenticator
    Auth.Setup "Mufasa", "Circle Of Life"
    Auth.Realm = "testrealm@host.com"
    Auth.ServerNonce = "dcd98b7102dd2f0e8b11d0f600bfb0c093"
    Auth.RequestCount = 1
    Auth.ClientNonce = "0a4f113b"
    Auth.Opaque = "5ccc069c403ebaf9f0171e9517f40e41"
    
    Dim Client As New WebClient
    Client.BaseUrl = "https://www.example.com/dir/"
    
    Dim Request As New WebRequest
    Request.Resource = "index.html"
    
    With Specs.It("should create header")
        Dim HA1 As String
        Dim HA2 As String
        Dim DigestResponse As String
        Dim ExpectedHeaders As String
        
        HA1 = "939e7578ed9e3c518a452acee763bce9"
        HA2 = "39aff3a2bab6126f332b942af96d3366"
        DigestResponse = WebHelpers.MD5(HA1 & ":" & Auth.ServerNonce & ":00000001:" & Auth.ClientNonce & ":auth:" & HA2)
        ExpectedHeaders = "Digest " & _
            "username=""Mufasa"", " & _
            "realm=""testrealm@host.com"", " & _
            "nonce=""" & Auth.ServerNonce & """, " & _
            "uri=""/dir/index.html"", " & _
            "qop=auth, " & _
            "nc=00000001, " & _
            "cnonce=""" & Auth.ClientNonce & """, " & _
            "response=""" & DigestResponse & """, " & _
            "opaque=""" & Auth.Opaque & """"
            
        .Expect(Auth.CreateHeader(Client, Request)).ToEqual ExpectedHeaders
    End With
    
    With Specs.It("should extract header information")
        Dim Unauthorized As New WebResponse
        Unauthorized.StatusCode = 401
        
        Unauthorized.Headers.Add WebHelpers.CreateKeyValue("WWW-Authenticate", "Digest realm=""testrealm@host.com""," & web_CrLf & _
                            "qop=""auth,auth-int""," & web_CrLf & _
                            "nonce=""dcd98b7102dd2f0e8b11d0f600bfb0c093""," & web_CrLf & _
                            "Opaque = ""5ccc069c403ebaf9f0171e9517f40e41""")
    
        Auth.Realm = ""
        Auth.ServerNonce = ""
        Auth.Opaque = ""
        
        Auth.ExtractAuthenticateInformation Unauthorized
        .Expect(Auth.Realm).ToEqual "testrealm@host.com"
        .Expect(Auth.ServerNonce).ToEqual "dcd98b7102dd2f0e8b11d0f600bfb0c093"
        .Expect(Auth.Opaque).ToEqual "5ccc069c403ebaf9f0171e9517f40e41"
    End With
    
    InlineRunner.RunSuite Specs
End Function


