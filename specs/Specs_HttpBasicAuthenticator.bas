Attribute VB_Name = "Specs_HttpBasicAuthenticator"
''
' Specs_HttpBasicAuthenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for HttpBasicAuthenticator
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "HttpBasicAuthenticator"
    
    Dim Auth As New HttpBasicAuthenticator
    
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Client.BaseUrl = HttpbinBaseUrl
    
    With Specs.It("should use Basic Authentication")
        Set Request = New WebRequest
        Request.Resource = "basic-auth/{user}/{password}"
        Request.AddUrlSegment "user", "Tim"
        Request.AddUrlSegment "password", "Secret123"
        
        Set Client.Authenticator = Nothing
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Unauthorized
        
        Auth.Setup "Tim", "Secret123"
        Set Client.Authenticator = Auth
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("authenticated")).ToEqual True
    End With
    
    With Specs.It("should properly escape username and password")
        Set Request = New WebRequest
        Request.Resource = "basic-auth/{user}/{password}"
        Request.AddUrlSegment "user", "Tim\`$""!"
        Request.AddUrlSegment "password", "Secret123\`$""!"
        
        Set Client.Authenticator = Nothing
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual WebStatusCode.Unauthorized
        
        Auth.Setup "Tim\`$""!", "Secret123\`$""!"
        Set Client.Authenticator = Auth
        
        Set Response = Client.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 200
        .Expect(Response.Data("authenticated")).ToEqual True
    End With
    
    InlineRunner.RunSuite Specs
End Function

