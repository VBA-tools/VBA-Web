VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SalesforceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private pConsumerSecret As String
Private pPassword As String
Private pSecurityToken As String

Public Property Get ConsumerKeyValue() As String
    If Me.[ConsumerKey].Value = "" And Credentials.Loaded Then
        ConsumerKeyValue = Credentials.Values("Salesforce")("key")
    Else
        ConsumerKeyValue = Me.[ConsumerKey].Value
    End If
End Property
Public Property Get ConsumerSecret() As String
    If pConsumerSecret = "" Then
        If Credentials.Loaded Then
            pConsumerSecret = Credentials.Values("Salesforce")("secret")
        Else
            pConsumerSecret = InputBox("Please enter Salesforce consumer secret")
        End If
    End If
    ConsumerSecret = pConsumerSecret
End Property
Public Property Get UsernameValue() As String
    If Me.[Username].Value = "" And Credentials.Loaded Then
        UsernameValue = Credentials.Values("Salesforce")("username")
    Else
        UsernameValue = Me.[Username].Value
    End If
End Property
Public Property Get Password() As String
    If pPassword = "" Then
        pPassword = InputBox("Please enter Salesforce password")
    End If
    Password = pPassword
End Property
Public Property Get SecurityToken() As String
    If pSecurityToken = "" Then
        If Credentials.Loaded Then
            pSecurityToken = Credentials.Values("Salesforce")("token")
        Else
            pSecurityToken = InputBox("Please enter Salesforce security token")
        End If
    End If
    SecurityToken = pSecurityToken
End Property

Public Sub GetObject()
    On Error GoTo ErrorHandling
    
    ClearOutput
    Login
'    If Me.[SalesforceUseAsync] Then
'        Salesforce.GetObjectAsync Me.[GetObjectType], Me.[GetObjectId], "SalesforceSheet.ProcessObject"
'    Else
        Dim Results As WebResponse
        Set Results = Salesforce.GetObject(Me.[GetObjectType], Me.[GetObjectId])
        
        ProcessObject Results
'    End If
    Exit Sub
    
ErrorHandling:
    
    OutputError Err.Number, Err.Description
    
End Sub

Public Sub ProcessObject(Results As WebResponse)
    If Results.StatusCode < 400 Then
        OutputObject Results.Data
    Else
        OutputError Results.StatusCode, Results.Content
    End If
End Sub

Private Sub ClearOutput()
    Me.[GetObjectResultName] = ""
End Sub

Private Sub OutputObject(Results As Dictionary)
    Me.[GetObjectResultName] = Results("Name")
End Sub

Private Sub OutputError(Code As Long, Message As String)
    Me.[GetObjectResultName] = "Error " & Code & ": " & Message
End Sub

Private Sub Login()
    Salesforce.ConsumerKey = ConsumerKeyValue
    Salesforce.ConsumerSecret = ConsumerSecret
    Salesforce.Username = UsernameValue
    Salesforce.Password = Password
    Salesforce.SecurityToken = SecurityToken
End Sub
