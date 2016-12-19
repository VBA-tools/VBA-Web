Attribute VB_Name = "Specs_WebAsyncWrapper"
''
' Specs_WebAsyncWrapper
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for WebAsyncWrapper
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Fixture As New AsyncFixture

Public Property Get HttpbinBaseUrl() As String
    HttpbinBaseUrl = "http://httpbin.org"
End Property

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebAsyncWrapper"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Specs
    Fixture.ListenTo Specs
    
    On Error Resume Next
    
    Dim Client As WebClient
    Dim Request As WebRequest
    Dim AsyncWrapper As WebAsyncWrapper
    Dim ClonedWrapper As WebAsyncWrapper
    
    Set Client = New WebClient
    Set AsyncWrapper = New WebAsyncWrapper
    Client.BaseUrl = HttpbinBaseUrl
    Set AsyncWrapper.Client = Client
    
    Dim WaitTime As Integer
    WaitTime = 500
    
    Dim SimpleCallback As String
    Dim ComplexCallback As String
    SimpleCallback = "Specs_WebAsyncWrapper.SimpleCallback"
    ComplexCallback = "Specs_WebAsyncWrapper.ComplexCallback"
    
    ' --------------------------------------------- '
    ' Properties
    ' --------------------------------------------- '
    
    ' Request
    ' Callback
    ' CallbackArgs
    ' Http
    ' Client
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ' ExecuteAsync
    ' --------------------------------------------- '
    With Specs.It("ExecuteAsync should pass response to callback")
        Set Request = New WebRequest
        Request.Resource = "get"
        
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait WaitTime * 2
        .Expect(Fixture.Response).ToNotBeUndefined
    End With
    
    With Specs.It("ExecuteAsync should pass arguments to callback")
        Set Request = New WebRequest
        Request.Resource = "get"
        
        AsyncWrapper.ExecuteAsync Request, ComplexCallback, Array("A", "B", "C")
        Wait WaitTime
        .Expect(Fixture.Response).ToNotBeUndefined
        If UBound(Fixture.Args) > 1 Then
            .Expect(Fixture.Args(0)).ToEqual "A"
            .Expect(Fixture.Args(1)).ToEqual "B"
            .Expect(Fixture.Args(2)).ToEqual "C"
        Else
            .Expect(UBound(Fixture.Args)).ToBeGreaterThan 1
        End If
    End With
    
    ' Clone
    ' @internal
    ' --------------------------------------------- '
    With Specs.It("Clone should copy Client")
        Set ClonedWrapper = AsyncWrapper.Clone
        
        .Expect(ClonedWrapper.Client).ToNotBeNothing
        .Expect(ClonedWrapper.Client.BaseUrl).ToEqual HttpbinBaseUrl
    End With
    
    ' PrepareAndExecuteRequest
    ' TimedOut
    
    ' Note: Weird async issues can occur if timeout spec isn't last
    With Specs.It("should return 408 and close request on request timeout")
        Set Request = New WebRequest
        Request.Resource = "delay/{seconds}"
        Request.AddUrlSegment "seconds", "2"

        Client.TimeoutMs = 100
        AsyncWrapper.ExecuteAsync Request, SimpleCallback
        Wait 2000
        .Expect(Fixture.Response).ToNotBeUndefined
        If Not Fixture.Response Is Nothing Then
            .Expect(Fixture.Response.StatusCode).ToEqual 408
            .Expect(Fixture.Response.StatusDescription).ToEqual "Request Timeout"
        End If
        .Expect(AsyncWrapper.Http).ToBeUndefined
        Client.TimeoutMs = 2000
    End With
End Function

Public Sub SimpleCallback(Response As WebResponse)
    Set Fixture.Response = Response
End Sub

Public Sub ComplexCallback(Response As WebResponse, Args As Variant)
    Set Fixture.Response = Response
    Fixture.Args = Args
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

