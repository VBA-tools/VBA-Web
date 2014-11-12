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

Public Property Get BaseUrl() As String
    BaseUrl = "http://httpbin.org"
End Property

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "Authenticator"
    
    Dim Client As New RestClient
    Dim Request As RestRequest
    Dim Response As RestResponse
    Dim Auth As SpecAuthenticator
    
    Client.BaseUrl = BaseUrl
    
    With Specs.It("should set parameter")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth

        Set Request = New RestRequest
        Request.Resource = "post"
        Request.AddBodyParameter "request_parameter", "request"
        Request.Method = httpPOST

        Set Response = Client.Execute(Request)
        .Expect(RestHelpers.ParseJSON(CStr(Response.Data("data")))("auth_parameter")).ToEqual "auth"
        .Expect(RestHelpers.ParseJSON(CStr(Response.Data("data")))("request_parameter")).ToEqual "request"
    End With
    
    With Specs.It("should set querystring parameter")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth
    
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddQuerystringParam "request_query", "request"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data("args")("auth_query")).ToEqual "auth"
        .Expect(Response.Data("args")("request_query")).ToEqual "request"
    End With
    
    With Specs.It("should set header")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth
        
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddHeader "X-Custom-B", "request"

        Set Response = Client.Execute(Request)
        .Expect(Response.Data("headers")("X-Custom-A")).ToEqual "auth"
        .Expect(Response.Data("headers")("X-Custom-B")).ToEqual "request"
    End With
    
    With Specs.It("should set cookie")
        Set Auth = New SpecAuthenticator
        Set Client.Authenticator = Auth

        Set Request = New RestRequest
        Request.Resource = "cookies"
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
        .Expect(Response.Data("headers")("Content-Type")).ToMatch "text/plain"
    End With
    
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function


