Attribute VB_Name = "Specs_DigestAuthenticator"
''
' Specs_DigestAuthenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for DigestAuthenticator
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "DigestAuthenticator"
    
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
    
    Dim Unauthorized As New WebResponse
    Unauthorized.StatusCode = 401
    
    Dim Header As New Dictionary
    Header.Add "key", "WWW-Authenticate"
    Header.Add "value", "Digest realm=""testrealm@host.com""," & vbCrLf & _
                        "qop=""auth,auth-int""," & vbCrLf & _
                        "nonce=""dcd98b7102dd2f0e8b11d0f600bfb0c093""," & vbCrLf & _
                        "Opaque = ""5ccc069c403ebaf9f0171e9517f40e41"""
    Unauthorized.Headers.Add Header
    
    Dim Expected As String
    Dim HA1 As String
    Dim HA2 As String
    Dim DigestResponse As String
    
    HA1 = "939e7578ed9e3c518a452acee763bce9"
    HA2 = "39aff3a2bab6126f332b942af96d3366"
    DigestResponse = WebHelpers.MD5(HA1 & ":" & Auth.ServerNonce & ":" & Auth.FormattedRequestCount & ":" & Auth.ClientNonce & ":auth:" & HA2)
    
    With Specs.It("should calculate HA1 from username, realm, and password")
        .Expect(Auth.CalculateHA1).ToEqual HA1
    End With
    
    With Specs.It("should calculate HA2 from method and uri")
        .Expect(Auth.CalculateHA2("GET", "/dir/index.html")).ToEqual HA2
    End With
    
    With Specs.It("should calculate response")
        .Expect(Auth.CalculateResponse(Client, Request)).ToEqual DigestResponse
    End With
    
    With Specs.It("should create header")
        Expected = "Digest " & _
            "username=""Mufasa"", " & _
            "realm=""testrealm@host.com"", " & _
            "nonce=""" & Auth.ServerNonce & """, " & _
            "uri=""/dir/index.html"", " & _
            "qop=auth, " & _
            "nc=" & Auth.FormattedRequestCount & ", " & _
            "cnonce=""" & Auth.ClientNonce & """, " & _
            "response=""" & DigestResponse & """, " & _
            "opaque=""" & Auth.Opaque & """"
            
        .Expect(Auth.CreateHeader(Client, Request)).ToEqual Expected
    End With
    
    With Specs.It("should extract header information")
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


