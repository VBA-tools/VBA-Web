VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwitterSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const TwitterResultsFirstRow As Integer = 4
Private Const TwitterResultsCount As Integer = 20
Private Const TwitterResultsCol As Integer = 2

Public Sub SearchTwitter()
    ClearResults
    
'    If Me.[UseAsync] Then
'        Twitter.SearchTwitterAsync Me.[TwitterQuery], "TwitterSheet.ProcessResults"
'    Else
        Dim Results As WebResponse
        Set Results = Twitter.SearchTwitter(Me.[TwitterQuery])
        
        ProcessResults Results
'    End If
End Sub

Private Sub ClearResults()
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim LastRow As Integer
    LastRow = TwitterResultsFirstRow + TwitterResultsCount - 1
    Me.Rows(TwitterResultsFirstRow & ":" & LastRow).ClearContents
    
    Application.ScreenUpdating = PrevUpdating
End Sub

Private Sub OutputResults(Results As Variant)
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim Status As Dictionary
    Dim Row As Integer
    
    Row = TwitterResultsFirstRow
    For Each Status In Results("statuses")
        Me.Cells(Row, TwitterResultsCol).Value = Status("text")
        
        Row = Row + 1
    Next Status
    
    Application.ScreenUpdating = PrevUpdating
End Sub

Private Sub OutputError(Code As Integer, Message As String)
    Me.Cells(TwitterResultsFirstRow, TwitterResultsCol) = "Error " & Code & ": " & Message
End Sub

Public Sub ProcessResults(Results As WebResponse)
    If Results.StatusCode < 400 Then
        OutputResults Results.Data
    Else
        OutputError Results.StatusCode, Results.Content
    End If
End Sub
