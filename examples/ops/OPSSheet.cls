VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OPSSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const OPSResultsFirstRow As Integer = 5
Private Const OPSResultsCount As Integer = 25
Private Const OPSResultsCol As Integer = 2

Public Sub SearchOPS()
    ClearResults
    
    Dim Results As Collection
    Dim Query As String
    
    Query = Me.[OPSQuery]
    If Query <> "" Then
        Set Results = OPS.Search(Query)
        OutputResults Results
    End If
End Sub

Private Sub ClearResults()
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim LastRow As Integer
    LastRow = OPSResultsFirstRow + OPSResultsCount - 1
    Me.Rows(OPSResultsFirstRow & ":" & LastRow).ClearContents
    
    Application.ScreenUpdating = PrevUpdating
End Sub

Private Sub OutputResults(Results As Collection)
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim Result As Dictionary
    Dim Row As Integer
    
    Row = OPSResultsFirstRow
    For Each Result In Results
        Me.Cells(Row, OPSResultsCol).Value = Result("title")
        
        Row = Row + 1
    Next Result
    
    Application.ScreenUpdating = PrevUpdating
End Sub

