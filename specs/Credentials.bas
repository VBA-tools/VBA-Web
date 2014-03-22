Attribute VB_Name = "Credentials"
Private Const CredentialsPath As String = "..\credentials.txt"
Private pCredentials As Dictionary

Public Property Get Values() As Dictionary
    If pCredentials Is Nothing Then
        Set pCredentials = Load
    End If
    Set Values = pCredentials
End Property

Public Property Get Loaded() As Boolean
    Loaded = Not Values Is Nothing
End Property
    

Function Load() As Dictionary
    Dim Line As String
    Dim Header As String
    Dim Parts As Variant
    Dim Key As String
    Dim Value As String
    
    Set pCredentials = New Dictionary
    Open FullPath(CredentialsPath) For Input As #1
    
    On Error GoTo ErrorHandling
    Do While Not EOF(1)
        Line Input #1, Line
        
        ' Skip blank lines and comment lines
        If Line <> "" And Left(Line, 1) <> "#" Then
            If Left(Line, 1) = "-" Then
                Line = Right(Line, Len(Line) - 1)
                Parts = Split(Line, ":")
                
                If UBound(Parts) >= 1 And Header <> "" And pCredentials.Exists(Header) Then
                    Key = Trim(Parts(0))
                    Value = Trim(Split(Parts(1), "#")(0))
                    
                    If Key <> "" And Value <> "" Then
                        pCredentials(Header).Add Key, Value
                    End If
                End If
            Else
                Header = Trim(Split(Line, "#")(0))
                
                If Header <> "" Then
                    pCredentials.Add Header, New Dictionary
                End If
            End If
        End If
    Loop
    
    Set Load = pCredentials
    
ErrorHandling:
    Close #1
End Function

Private Function FullPath(RelativePath As String) As String
    FullPath = ThisWorkbook.Path & "\" & RelativePath
End Function
