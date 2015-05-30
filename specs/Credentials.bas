Attribute VB_Name = "Credentials"
Private pCredentials As Dictionary

Public Property Get CredentialsPath() As String
    ' Go up one folder from workbook path
    Dim Parts() As String
    Dim i As Long
    Parts = VBA.Split(ThisWorkbook.Path, Application.PathSeparator)
    For i = LBound(Parts) To UBound(Parts) - 1
        If CredentialsPath = "" Then
            CredentialsPath = CredentialsPath & Parts(i)
        Else
            CredentialsPath = CredentialsPath & Application.PathSeparator & Parts(i)
        End If
    Next i
    
    CredentialsPath = CredentialsPath & Application.PathSeparator & "credentials.txt"
End Property

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
    Open CredentialsPath For Input As #1
    
    On Error GoTo ErrorHandling
    Do While Not VBA.EOF(1)
        Line Input #1, Line
        Line = VBA.Replace(Line, vbNewLine, "")
        
        ' Skip blank lines and comment lines
        If Line <> "" And VBA.Left$(Line, 1) <> "#" Then
            If VBA.Left$(Line, 1) = "-" Then
                Line = VBA.Right$(Line, VBA.Len(Line) - 1)
                Parts = VBA.Split(Line, ":", 2)
                
                If UBound(Parts) >= 1 And Header <> "" And pCredentials.Exists(Header) Then
                    Key = VBA.Trim(Parts(0))
                    Value = VBA.Trim(Split(Parts(1), "#")(0))
                    
                    If Key <> "" And Value <> "" Then
                        pCredentials(Header).Add Key, Value
                    End If
                End If
            Else
                Header = VBA.Trim(VBA.Split(Line, "#")(0))
                
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
