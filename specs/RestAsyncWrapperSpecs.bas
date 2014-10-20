Attribute VB_Name = "RestAsyncWrapperSpecs"
''
' RestAsyncWrapperSpecs
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
    Specs.Description = "RestAsyncWrapper"
    Specs.BeforeEach "RestAsyncWrapperSpecs.Reset"
    
    On Error Resume Next
    
    Dim Client As New RestClient
    Client.BaseUrl = "http://localhost:3000"
    
    Dim Request As RestRequest
    Dim AsyncWrapper As New RestAsyncWrapper
    Set AsyncWrapper.Client = Client
    
    Dim WaitTime As Integer
    WaitTime = 500
    
    Dim SimpleCallback As String
    Dim ComplexCallback As String
    SimpleCallback = "RestAsyncWrapperSpecs.SimpleCallback"
    ComplexCallback = "RestAsyncWrapperSpecs.ComplexCallback"
    
    Dim BodyToString As String
    
    With Specs.It("should pass response to callback")
        Set Request = New RestRequest
        Request.Resource = "get"
        
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime * 2
        .Expect(AsyncResponse).ToBeDefined
    End With
    
    With Specs.It("should pass arguments to callback")
        Set Request = New RestRequest
        Request.Resource = "get"
        
        AsyncWrapper.ExecuteAsync Request, ComplexCallback, Array("A", "B", "C")
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
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 200
        .Expect(AsyncResponse.StatusDescription).ToEqual "OK"
        
        Request.AddUrlSegment "code", 304
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 304
        .Expect(AsyncResponse.StatusDescription).ToEqual "Not Modified"
        
        Request.AddUrlSegment "code", 404
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 404
        .Expect(AsyncResponse.StatusDescription).ToEqual "Not Found"
        
        Request.AddUrlSegment "code", 500
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse.StatusCode).ToEqual 500
        .Expect(AsyncResponse.StatusDescription).ToEqual "Internal Server Error"
    End With
    
    With Specs.It("should include binary body in response")
        Set Request = New RestRequest
        Request.Resource = "howdy"
        
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse).ToBeDefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.Body).ToBeDefined
            
            If Not IsEmpty(AsyncResponse.Body) Then
                For i = LBound(AsyncResponse.Body) To UBound(AsyncResponse.Body)
                    BodyToString = BodyToString & Chr(AsyncResponse.Body(i))
                Next i
            End If
            
            .Expect(BodyToString).ToEqual "Howdy!"
        End If
    End With
    
    With Specs.It("should include headers in response")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse).ToBeDefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.Headers.Count).ToBeGTE 5
        
            Dim Header As Dictionary
            Dim NumCookies As Integer
            For Each Header In AsyncResponse.Headers
                If Header("key") = "Set-Cookie" Then
                    NumCookies = NumCookies + 1
                End If
            Next Header
            
            .Expect(NumCookies).ToEqual 5
        End If
    End With
    
    With Specs.It("should include cookies in response")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse).ToBeDefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.Cookies.Count).ToEqual 4
            .Expect(AsyncResponse.Cookies("unsigned-cookie")).ToEqual "simple-cookie"
            .Expect(AsyncResponse.Cookies("signed-cookie")).ToContain "special-cookie"
            .Expect(AsyncResponse.Cookies("tricky;cookie")).ToEqual "includes; semi-colon and space at end "
            .Expect(AsyncResponse.Cookies("duplicate-cookie")).ToEqual "B"
        End If
    End With
    
    With Specs.It("should include cookies with request")
        Set Request = New RestRequest
        Request.Resource = "cookie"
        
        Set Response = Client.Execute(Request)
    
        Set Request = New RestRequest
        Request.Resource = "get"
        Request.AddCookie "test-cookie", "howdy"
        Request.AddCookie "signed-cookie", Response.Cookies("signed-cookie")
        
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime
        .Expect(AsyncResponse).ToBeDefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.Data).ToBeDefined
            If Not IsEmpty(AsyncResponse.Data) Then
                .Expect(AsyncResponse.Data("cookies").Count).ToEqual 1
                .Expect(AsyncResponse.Data("cookies")("test-cookie")).ToEqual "howdy"
                .Expect(AsyncResponse.Data("signed_cookies").Count).ToEqual 1
                .Expect(AsyncResponse.Data("signed_cookies")("signed-cookie")).ToEqual "special-cookie"
            End If
        End If
    End With
    
    ' Note: Weird async issues can occur if timeout spec isn't last
    With Specs.It("should return 408 and close request on request timeout")
        Set Request = New RestRequest
        Request.Resource = "timeout"
        Request.AddQuerystringParam "ms", 2000

        Client.TimeoutMS = 100
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait 2000
        .Expect(AsyncResponse).ToBeDefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.StatusCode).ToEqual 408
            .Expect(AsyncResponse.StatusDescription).ToEqual "Request Timeout"
        End If
        .Expect(AsyncWrapper.Http).ToBeUndefined
        Client.TimeoutMS = 2000
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

