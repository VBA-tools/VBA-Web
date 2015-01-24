Attribute VB_Name = "Specs_IWebAuthenticator"
''
' Specs_IWebAuthenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for IWebAuthenticator
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "IWebAuthenticator"
    
    Dim Client As New WebClient
    Dim Request As WebRequest
    Dim Response As WebResponse
    Dim Auth As SpecAuthenticator
    
    Client.BaseUrl = HttpbinBaseUrl
    
'    With Specs.It("should set body, querystring, headers, cookies")
'        Set Auth = New SpecAuthenticator
'        Set Client.Authenticator = Auth
'
'        Set Request = New WebRequest
'        Request.Resource = "get"
'        Request.AddBodyParameter "request_parameter", "request"
'        Request.Method = httpGET
'
'        Set Response = Client.Execute(Request)
'        .Expect(WebHelpers.ParseJSON(CStr(Response.Data("data")))("auth_parameter")).ToEqual "auth"
'        .Expect(WebHelpers.ParseJSON(CStr(Response.Data("data")))("request_parameter")).ToEqual "request"
'    End With

'    With Specs.It("should set querystring parameter")
'        Set Auth = New SpecAuthenticator
'        Set Client.Authenticator = Auth
'
'        Set Request = New WebRequest
'        Request.Resource = "get"
'        Request.AddQuerystringParam "request_query", "request"
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Data("args")("auth_query")).ToEqual "auth"
'        .Expect(Response.Data("args")("request_query")).ToEqual "request"
'    End With
'
'    With Specs.It("should set header")
'        Set Auth = New SpecAuthenticator
'        Set Client.Authenticator = Auth
'
'        Set Request = New WebRequest
'        Request.Resource = "get"
'        Request.AddHeader "X-Custom-B", "request"
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Data("headers")("X-Custom-A")).ToEqual "auth"
'        .Expect(Response.Data("headers")("X-Custom-B")).ToEqual "request"
'    End With
'
'    With Specs.It("should set cookie")
'        Set Auth = New SpecAuthenticator
'        Set Client.Authenticator = Auth
'
'        Set Request = New WebRequest
'        Request.Resource = "cookies"
'        Request.AddCookie "request_cookie", "request"
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Data("cookies")("auth_cookie")).ToEqual "auth"
'        .Expect(Response.Data("cookies")("request_cookie")).ToEqual "request"
'    End With
'
'    With Specs.It("should set content-type")
'        Set Auth = New SpecAuthenticator
'        Set Client.Authenticator = Auth
'
'        Set Request = New WebRequest
'        Request.Resource = "get"
'
'        Set Response = Client.Execute(Request)
'        .Expect(Response.Data("headers")("Content-Type")).ToMatch "text/plain"
'    End With
    
    Set Client = Nothing
    
    InlineRunner.RunSuite Specs
End Function


