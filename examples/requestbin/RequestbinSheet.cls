Option Explicit

Private Const RequestbinResultsFirstRow As Integer = 2      'Row 2
Private Const RequestbinResultsCol As Integer = 2           'Column B
Private Const RequestbinResultsCount As Integer = 1

Public Sub SearchRequestbin()
    Dim Response As WebResponse
    Dim Tempurl As String, Post_data As String              'declared variables for readability
    Tempurl = Range("E5")
    Post_data = Range("B1")

    ClearResults                                            'calls ClearResults()
    WebHelpers.EnableLogging = True                         'extended debug info
    
    If IsEmpty(Tempurl) = True Then                         'user must enter their personal requestbin url
        MsgBox ("Please enter your request bin url")
        Exit Sub
    End If
    
    If IsEmpty(Post_data) = False Then                      'don't submit a blank post data
        Set Response = RequestbinLookup(Tempurl, Post_data) 'this line goes out and does everything
    Else                                                    'when the program comes back to here it is already finished running
        MsgBox ("Input is empty")
        Exit Sub
    End If

    ProcessResults Response                                 'calls ProcessResults() below with Response, the webresponse we received
End Sub

Public Sub ProcessResults(Results As WebResponse)
    If Results.StatusCode < 400 Then
        OutputResults Results                               'calls OutputResults()
    Else
        OutputError Results.StatusCode, Results.Content     'calls OutputError()
    End If
End Sub

Private Sub OutputResults(Results As WebResponse)
    'requestbin just lets you post data, it only returns a server response 'ok'
    Me.Cells(RequestbinResultsFirstRow, RequestbinResultsCol) = Results.Content
End Sub

Private Sub OutputError(Code As Integer, Message As String)
    Me.Cells(RequestbinResultsFirstRow, RequestbinResultsCol) = "Error " & Code & ": " & Message
End Sub

Private Sub ClearResults()
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim LastRow As Integer
    LastRow = RequestbinResultsFirstRow + RequestbinResultsCount - 1
    'Me.Rows(RequestbinResultsFirstRow & ":" & LastRow).ClearContents       'clear entire row
    Me.Range(Me.Cells(RequestbinResultsFirstRow, RequestbinResultsCol), Me.Cells(LastRow, RequestbinResultsCol)).ClearContents           'clear selected part of column
    
    Application.ScreenUpdating = PrevUpdating
End Sub
