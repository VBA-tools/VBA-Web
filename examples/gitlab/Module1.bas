Public Const Project_CoreGrid = 6452557
Public Const Project_Jhipster = 6822181
Public Const Project_Audit = 7277583

Sub Button1_Click()
    
    
    Dim PrevUpdating As Boolean
    Dim LastIndex As Integer
    
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    LastIndex = 1
    GetIssues Project_CoreGrid, LastIndex
    GetIssues Project_Jhipster, LastIndex
    GetIssues Project_Audit, LastIndex
    FindStartDates
    'GetEvents
    'ActiveSheet.Cells(5, 1) = "Yukleniyor
    'ActiveSheet.Cells(5, 1) = gitlab.GetIssiues
    
    Application.ScreenUpdating = PrevUpdating
End Sub

Sub GetIssues(ProjectId As Long, Optional ByRef LastIndex As Integer = 1, Optional LastPage As Integer = 1)
    Dim Issues As Object
    Dim Assignee As Object
    Dim Sheet As Worksheet
    
    Set Issues = gitlab.GetIssiues(ProjectId, LastPage)
    
    Set Sheet = Worksheets("issues")
    If (LastIndex = 1) Then
        Sheet.UsedRange.Clear
        Sheet.Cells(1, 1) = "project_id"
        Sheet.Cells(1, 2) = "id"
        Sheet.Cells(1, 3) = "iid"
        Sheet.Cells(1, 4) = "title"
        Sheet.Cells(1, 5) = "state"
        Sheet.Cells(1, 6) = "assignee.name"
        Sheet.Cells(1, 7) = "created_at"
        Sheet.Cells(1, 8) = "closed_at"
        LastIndex = LastIndex + 1
    End If
    For Each Issue In Issues
        Sheet.Cells(LastIndex, 1) = ProjectId
        Sheet.Cells(LastIndex, 2) = Issue("id")
        Sheet.Cells(LastIndex, 3) = Issue("iid")
        Sheet.Cells(LastIndex, 4) = Issue("title")
        Sheet.Cells(LastIndex, 5) = Issue("state")
        If Not IsNull(Issue("assignee")) Then
            Set Assignee = Issue("assignee")
            Sheet.Cells(LastIndex, 6) = Assignee("name")
        End If
        Sheet.Cells(LastIndex, 7) = FormatDate(Issue("created_at"))
        Sheet.Cells(LastIndex, 8) = FormatDate(Issue("closed_at"))
        LastIndex = LastIndex + 1
    Next Issue
    If LastIndex = LastPage * 100 + 2 Then
        GetIssues ProjectId, LastIndex, LastPage + 1
    End If
End Sub

Sub FindStartDates()
    Dim Notes As Object
    Dim Assignee As Object
    Dim Sheet As Worksheet
    Dim ProjectId As Long
    Dim IssueId As Integer
    
    Set Sheet = Worksheets("issues")
    Sheet.Cells(1, 9) = "started_at"
    For Each Row In Sheet.Rows
        'Bos satira gelince duralim
        If IsEmpty(Sheet.Cells(Row.Row, 1).value) Then
          Exit For
        End If
        'ilk satırı geçelim
        If Row.Row > 1 Then
            ProjectId = Sheet.Cells(Row.Row, 1).value
            IssueId = Sheet.Cells(Row.Row, 3).value
            Set Notes = gitlab.GetNotes(ProjectId, IssueId)
            If Not Notes Is Nothing Then
                For Each Note In Notes
                    If InStrRev(Note("body"), "added ~4112781 label", -1, vbTextCompare) > 0 Then
                        Sheet.Cells(Row.Row, 9) = FormatDate(Note("created_at"))
                        Exit For
                    End If
                Next Note
            End If
        End If
    Next Row
End Sub

Sub GetEvents()
    Dim Events As Object
    Dim Assignee As Object
    Dim Sheet As Worksheet
    Dim i As Integer
    
    Set Events = gitlab.GetEvents
    
    Set Sheet = Worksheets("events")
    Sheet.UsedRange.Clear
    Sheet.Cells(1, 1) = "issue_id"
    Sheet.Cells(1, 2) = "action_name"
    Sheet.Cells(1, 3) = "created_at"
    i = 2
    For Each Evnt In Events
        Sheet.Cells(i, 1) = Evnt("target_id")
        Sheet.Cells(i, 2) = Evnt("action_name")
        Sheet.Cells(i, 3) = Evnt("created_at")
        i = i + 1
    Next Evnt
    
End Sub

Function FormatDate(value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        FormatDate = ""
    Else
        FormatDate = Mid(value, 9, 2) & "." & Mid(value, 6, 2) & ". " & Mid(value, 1, 4) & " " & Mid(value, 12, 8)
    End If
End Function
