Attribute VB_Name = "RestClientBaseSpecs"
''
' RestClientAsyncSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Async specs for the RestRequest class
'
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim AsyncResponses As Collection
Dim AsyncArgs As Collection

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "RestClientBase"
    Specs.BeforeEach "RestclientBaseSpecs.Reset"
    
    Dim SimpleCallback As String
    SimpleCallback = "RestClientBaseSpecs.SimpleCallback"
    
    Dim Request As RestRequest
    Dim SecondRequest As RestRequest
    Dim Response As RestResponse

    RestClientBase.BaseUrl = "localhost:3000"
        
    With Specs.It("should perform sync requests")
        Set Request = New RestRequest
        
        Request.Resource = "status/304"
        Set Response = RestClientBase.Execute(Request)
        .Expect(Response.StatusCode).ToEqual 304
        
        Request.Resource = "get"
        Request.AddQuerystringParam "q", "Howdy!"
        Set Response = RestClientBase.Execute(Request)
        .Expect(Response.Data("query")("q")).ToEqual "Howdy!"
        
        Request.Resource = "text"
        Request.ContentType = "text/plain"
        Request.Method = httpPOST
        Request.AddBodyString "Inline body"
        Set Response = RestClientBase.Execute(Request)
        .Expect(Response.Data("body")).ToEqual "Inline body"
    End With
    
    With Specs.It("should execute multiple async requests simultaneously")
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddQuerystringParam "a", 123
        
        Set SecondRequest = New RestRequest
        SecondRequest.Resource = "get"
        SecondRequest.AddQuerystringParam "b", 456
        
        RestClientBase.ExecuteAsync Request, SimpleCallback
        Wait 100
        RestClientBase.ExecuteAsync SecondRequest, SimpleCallback
        
        Wait 1000
        .Expect(AsyncResponses.count).ToEqual 2
        .Expect(AsyncResponses(1).Data("query")("a")).ToEqual "123"
        .Expect(AsyncResponses(2).Data("query")("b")).ToEqual "456"
    End With
    
    InlineRunner.RunSuite Specs
End Function

Sub SimpleCallback(Response As RestResponse)
    AsyncResponses.Add Response
End Sub

Sub Reset()
    Set AsyncResponses = New Collection
    Set AsyncArgs = New Collection
End Sub

Public Sub Wait(Milliseconds As Integer)
    Dim EndTime As Long
    EndTime = GetTickCount() + Milliseconds
    
    Dim WaitInterval As Long
    WaitInterval = Ceiling(Milliseconds / Ceiling(Milliseconds / 100))
    
    Do
        Sleep WaitInterval
        DoEvents
    Loop Until GetTickCount() >= EndTime
End Sub

Private Function Ceiling(Value As Double) As Long
    Ceiling = Int(Value)
    If Ceiling < Value Then
        Ceiling = Ceiling + 1
    End If
End Function

