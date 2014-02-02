Attribute VB_Name = "RestClientAsyncSpecs"
''
' RestClientAsyncSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Async specs for the RestRequest class
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim AsyncResponse As RestResponse
Dim AsyncArgs As Variant

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "RestClient Async"
    Specs.BeforeEach "RestClientAsyncSpecs.Reset"
    
    Dim Client As New RestClient
    Client.BaseUrl = "http://localhost:3000"
    
    Dim Request As RestRequest
    Dim WaitTime As Integer
    WaitTime = 500
    
    Dim SimpleCallback As String
    Dim ComplexCallback As String
    SimpleCallback = "RestClientAsyncSpecs.SimpleCallback"
    ComplexCallback = "RestClientAsyncSpecs.ComplexCallback"
    
    With Specs.It("should pass response to callback")
        Set Request = New RestRequest
        Request.Resource = "get"
        
        Client.ExecuteAsync Request, SimpleCallback
        Wait WaitTime * 2
        .Expect(AsyncResponse).ToBeDefined
    End With
    
    With Specs.It("should pass arguments to callback")
        Set Request = New RestRequest
        Request.Resource = "get"
        
        Client.ExecuteAsync Request, ComplexCallback, Array("A", "B", "C")
        Wait WaitTime
        .Expect(AsyncResponse).ToBeDefined
        If UBound(AsyncArgs) > 1 Then
            .Expect(AsyncArgs(0)).ToEqual "A"
            .Expect(AsyncArgs(1)).ToEqual "B"
            .Expect(AsyncArgs(2)).ToEqual "C"
        Else
            .Expect(UBound(AsyncArgs)).ToBeGreaterThan 1
        End If
    End With
    
    With Specs.It("should pass status and status description to response")
        Set Request = New RestRequest
        Request.Resource = "status/{code}"
        
        Request.AddUrlSegment "code", 200
        Client.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 200
        .Expect(AsyncResponse.StatusDescription).ToEqual "OK"
        
        Request.AddUrlSegment "code", 304
        Client.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 304
        .Expect(AsyncResponse.StatusDescription).ToEqual "Not Modified"
        
        Request.AddUrlSegment "code", 404
        Client.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 404
        .Expect(AsyncResponse.StatusDescription).ToEqual "Not Found"
        
        Request.AddUrlSegment "code", 500
        Client.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 500
        .Expect(AsyncResponse.StatusDescription).ToEqual "Internal Server Error"
    End With
    
    With Specs.It("should return 408 and close request on request timeout")
        Set Request = New RestRequest
        Request.Resource = "timeout"
        Request.AddQuerystringParam "ms", 2000

        Client.TimeoutMS = 100
        Client.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse).ToBeDefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.StatusCode).ToEqual 408
            .Expect(AsyncResponse.StatusDescription).ToEqual "Request Timeout"
        End If
        .Expect(Request.HttpRequest).ToBeUndefined
    End With
    
    InlineRunner.RunSuite Specs
End Function

Public Sub SimpleCallback(Response As RestResponse)
    Set AsyncResponse = Response
End Sub

Public Sub ComplexCallback(Response As RestResponse, Args As Variant)
    Set AsyncResponse = Response
    AsyncArgs = Args
End Sub

Public Sub Reset()
    Set AsyncResponse = New RestResponse
    AsyncArgs = Array()
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

