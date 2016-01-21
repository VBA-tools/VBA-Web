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
Dim AsyncResponse As WebResponse
Dim AsyncArgs As Variant

Public Property Get HttpbinBaseUrl() As String
    HttpbinBaseUrl = "http://httpbin.org"
End Property

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "WebAsyncWrapper"
    Specs.BeforeEach "Specs_WebAsyncWrapper.Reset"
    
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
        .Expect(AsyncResponse).ToNotBeUndefined
    End With
    
    With Specs.It("ExecuteAsync should pass arguments to callback")
        Set Request = New WebRequest
        Request.Resource = "get"
        
        AsyncWrapper.ExecuteAsync Request, ComplexCallback, Array("A", "B", "C")
        Wait WaitTime
        .Expect(AsyncResponse).ToNotBeUndefined
        If UBound(AsyncArgs) > 1 Then
            .Expect(AsyncArgs(0)).ToEqual "A"
            .Expect(AsyncArgs(1)).ToEqual "B"
            .Expect(AsyncArgs(2)).ToEqual "C"
        Else
            .Expect(UBound(AsyncArgs)).ToBeGreaterThan 1
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
        .Expect(AsyncResponse).ToNotBeUndefined
        If Not AsyncResponse Is Nothing Then
            .Expect(AsyncResponse.StatusCode).ToEqual 408
            .Expect(AsyncResponse.StatusDescription).ToEqual "Request Timeout"
        End If
        .Expect(AsyncWrapper.Http).ToBeUndefined
        Client.TimeoutMs = 2000
    End With
    
    InlineRunner.RunSuite Specs
End Function

Public Sub SimpleCallback(Response As WebResponse)
    Set AsyncResponse = Response
End Sub

Public Sub ComplexCallback(Response As WebResponse, Args As Variant)
    Set AsyncResponse = Response
    AsyncArgs = Args
End Sub

Public Sub Reset()
    Set AsyncResponse = New WebResponse
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

