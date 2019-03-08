Attribute VB_Name = "Module1"
Sub Button1_Click()
    
    Dim PrevUpdating As Boolean
    Dim LastIndex As Integer
    
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    LastIndex = 1
    GetIssues ActiveSheet.Cells(1, 2), LastIndex
    
    Application.ScreenUpdating = PrevUpdating
    
End Sub

Sub Button2_Click()
    
    Dim PrevUpdating As Boolean
    Dim LastIndex As Integer
    
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    LastIndex = 1
    GetProjects
    
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

Sub GetProjects()
    Dim Projects As Object
    Dim Assignee As Object
    Dim Sheet As Worksheet
    Dim LastIndex As Integer
    
    Set Projects = gitlab.GetProjects()
    
    Set Sheet = Worksheets("projects")
    Sheet.UsedRange.Clear
    Sheet.Cells(1, 1) = "id"
    Sheet.Cells(1, 2) = "name"

   LastIndex = 2
    For Each Project In Projects
        Sheet.Cells(LastIndex, 1) = Project("id")
        Sheet.Cells(LastIndex, 2) = Project("name")
        LastIndex = LastIndex + 1
    Next Project
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

Function Decode(value As String) As String
    value = Replace(value, "ş", "s")
    value = Replace(value, "ı", "i")
    value = Replace(value, "İ", "I")
    Decode = value
End Function

Function FormatDate(value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        FormatDate = ""
    Else
        FormatDate = Mid(value, 9, 2) & "." & Mid(value, 6, 2) & ". " & Mid(value, 1, 4) & " " & Mid(value, 12, 8)
    End If
End Function



