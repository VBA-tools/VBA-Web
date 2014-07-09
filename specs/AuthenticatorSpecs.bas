Attribute VB_Name = "AuthenticatorSpecs"
''
' RestClientSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' General and sync specs for the RestClient class
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "Authenticator"
    
    Dim Client As New RestClient
    Dim Request As RestRequest
    Dim Response As RestResponse
    Dim Auth As SpecAuthenticator
    
    Client.BaseUrl = "http://localhost:3000/"
    
    With Specs.It("should set parameter")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth

        Set Request = New RestRequest
        Request.Resource = "post"
        Request.AddParameter "request_parameter", "request"
        Request.Method = httpPOST

        Set Response = Client.Execute(Request)
        .Expect(RestHelpers.ParseJSON(CStr(Response.Data("body")))("auth_parameter")).ToEqual "auth"
        .Expect(RestHelpers.ParseJSON(CStr(Response.Data("body")))("request_parameter")).ToEqual "request"
    End With
    
    With Specs.It("should set querystring parameter")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth
    
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddQuerystringParam "request_query", "request"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data("query")("auth_query")).ToEqual "auth"
        .Expect(Response.Data("query")("request_query")).ToEqual "request"
    End With
    
    With Specs.It("should set header")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth
        
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddHeader "custom-b", "request"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data("headers")("custom-a")).ToEqual "auth"
        .Expect(Response.Data("headers")("custom-b")).ToEqual "request"
    End With
    
    With Specs.It("should set cookie")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth
        
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddCookie "request_cookie", "request"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data("cookies")("auth_cookie")).ToEqual "auth"
        .Expect(Response.Data("cookies")("request_cookie")).ToEqual "request"
    End With
    
    With Specs.It("should set content-type")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth
        
        Set Request = New RestRequest
        Request.Resource = "get"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data("headers")("content-type")).ToEqual "text/plain"
    End With
    
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function


