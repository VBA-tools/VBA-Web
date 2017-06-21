Option Explicit

Private Const HttpbinResultsFirstRow As Integer = 3     'Row 3
Private Const HttpbinResultsCol As Integer = 2          'Column B
Private Const HttpbinResultsCount As Integer = 6


Public Sub SearchHttpBin()
    Dim Response As WebResponse
    Dim RequestUrl As String, Post_data As String               'declared variables for readability
    RequestUrl = LCase(Range("B1"))
    Post_data = Range("B2")

    ClearResults                                                'calls ClearResults()
    WebHelpers.EnableLogging = True                             'extended debug info

    If RequestUrl = "post" Then                                 'if posting, make sure there is post data in RequestURL
        If Post_data = "" Then
            MsgBox ("Post Data input is empty")
            Exit Sub
        End If
    End If
    
    If RequestUrl <> "" Then                                    'make sure cell B1 has data
        Set Response = HttpbinLookup(RequestUrl, Post_data)     'call HttpbinLookup in Module Httpbin
    Else
        MsgBox ("Request input is empty")                       'when the program comes back to here it is finished running
        Exit Sub
    End If

    ProcessResults Response                                     'calls ProcessResults() below with Response, the webresponse we received
End Sub


Public Sub ProcessResults(Results As WebResponse)
    If Results.StatusCode < 400 Then
        OutputResults Results                                   'calls OutputResults()
    Else
        OutputError Results.StatusCode, Results.Content         'calls OutputError()
    End If
End Sub


Private Sub OutputResults(Results As WebResponse)
    Dim request As String
    request = LCase(Range("B1"))
    If request = "get?show_env=1" Then
        Range("B3") = "url: " & Results.Data("url")
        Range("B4") = "user-agent: " & Results.Data("headers")("User-Agent")
        Range("B5") = "origin: " & Results.Data("origin")
        Range("B6") = "protocol: " & Results.Data("headers")("X-Forwarded-Proto")
        Range("B7") = "port: " & Results.Data("headers")("X-Forwarded-Port")
    ElseIf request = "get" Then
        Range("B3") = "url: " & Results.Data("url")
        Range("B4") = "user-agent: " & Results.Data("headers")("User-Agent")
        Range("B5") = "origin: " & Results.Data("origin")
    ElseIf Left(request, 10) = "basic-auth" Then
        Range("B3") = "authenticated: " & Results.Data("authenticated")
        Range("B4") = "user: " & Results.Data("user")
    ElseIf Left(request, 11) = "digest-auth" Then
        Range("B3") = "authenticated: " & Results.Data("authenticated")
        Range("B4") = "user: " & Results.Data("user")
    ElseIf request = "post" Then
        'these will give you an error if they are not returned in the response (because you requested something different)
        Range("B3") = Results.Data("data")                      'outputs all data in one line, probably not what you want
        Range("B4") = "url: " & Results.Data("url")
        Range("B5") = "user-agent: " & Results.Data("headers")("User-Agent")
        Range("B6") = "origin: " & Results.Data("origin")
        'don't try reading individual entries out of ("data")
        'for it to work they all need to be on separate lines in the debug window, like how "headers" or "json" are
        Range("B7") = "systemtime: " & Results.Data("json")("systemtime")
        Range("B8") = "postdata: " & Results.Data("json")("postdata")
    End If
End Sub


Private Sub OutputError(Code As Integer, Message As String)
    Me.Cells(HttpbinResultsFirstRow, HttpbinResultsCol) = "Error " & Code & ": " & Message
End Sub


Private Sub ClearResults()
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim LastRow As Integer
    LastRow = HttpbinResultsFirstRow + HttpbinResultsCount - 1
    'Me.Rows(HttpbinResultsFirstRow & ":" & LastRow).ClearContents          'clear entire row
    Me.Range(Me.Cells(HttpbinResultsFirstRow, HttpbinResultsCol), Me.Cells(LastRow, HttpbinResultsCol)).ClearContents           'clear selected part of column
    
    Application.ScreenUpdating = PrevUpdating
End Sub
