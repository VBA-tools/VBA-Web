Attribute VB_Name = "Salesforce"
' Salesforce client with RESTful methods
' http://www.salesforce.com/us/developer/docs/api_rest/index.htm

Private Const ApiVersion As String = "v26.0"
Private pClient As WebClient

Public ConsumerKey As String
Public ConsumerSecret As String
Public Username As String
Public Password As String
Public SecurityToken As String

Public Property Get Client() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "https://na15.salesforce.com/"
        
        ' Setup OAuth2 authentication
        Dim Auth As New OAuth2Authenticator
        Auth.Setup _
            ClientId:=ConsumerKey, _
            ClientSecret:=ConsumerSecret, _
            Username:=Username, _
            Password:=Password & SecurityToken
        
        Auth.TokenUrl = "https://login.salesforce.com/services/oauth2/token"
        Set pClient.Authenticator = Auth
    End If
    
    Set Client = pClient
End Property

''
' Get generic object (async and sync)
' --------------------------------------- '
'Public Sub GetObjectAsync(ObjectName As String, ObjectId As String, _
'    Callback As String, ParamArray CallbackArgs() As Variant)
'
'    Call Client.ExecuteAsync(ObjectRequest(ObjectName, ObjectId), Callback, CallbackArgs)
'End Sub

Public Function GetObject(ObjectName As String, ObjectId As String) As WebResponse
    Set GetObject = Client.Execute(ObjectRequest(ObjectName, ObjectId))
End Function

''
' Update generic object (async and sync)
' --------------------------------------- '
'Public Sub UpdateObjectAsync(ObjectName As String, ObjectId As String, Values As Dictionary, _
'    Callback As String, ParamArray CallbackArgs() As Variant)
'
'    Call Client.ExecuteAsync(UpdateRequest(ObjectName, ObjectId, Values), Callback, CallbackArgs)
'End Sub

Public Function UpdateObject(ObjectName As String, ObjectId As String, Values As Dictionary) As WebResponse
    Set UpdateObject = Client.Execute(UpdateRequest(ObjectName, ObjectId, Values))
End Function

''
' Run query (async and sync)
' --------------------------------------- '
'Public Sub RunQueryAsync(query As String, _
'    Callback As String, ParamArray CallbackArgs() As Variant)
'
'    Call Client.ExecuteAsync(QueryRequest(query), Callback, CallbackArgs)
'End Sub

Public Function RunQuery(query As String) As WebResponse
    Set RunQuery = Client.Execute(QueryRequest(query))
End Function

''
' Get overview info for the specified object
' --------------------------------------- '
Public Function GetObjectInfo(ObjectName As String) As WebResponse
    Dim Request As WebRequest
    Set Request = ObjectRequest(ObjectName, "describe")
    
    Set GetObjectInfo = Client.Execute(Request)
End Function

''
' Setup generic object request
' --------------------------------------- '
Public Function ObjectRequest(ObjectName As String, ObjectId As String) As WebRequest
    Dim Request As New WebRequest
    Request.Resource = "services/data/{ApiVersion}/sobjects/{ObjectName}/{ObjectId}"
    Request.AddUrlSegment "ApiVersion", ApiVersion
    Request.AddUrlSegment "ObjectName", ObjectName
    Request.AddUrlSegment "ObjectId", ObjectId
    Set ObjectRequest = Request
End Function

''
' Setup generic update request
' --------------------------------------- '
Public Function UpdateRequest(ObjectName As String, ObjectId As String, Values As Dictionary) As WebRequest
    Dim Request As WebRequest
    Set Request = ObjectRequest(ObjectName, ObjectId)
    
    ' Remove Id from values for update
    If Values.Exists("Id") Then Values.Remove "Id"
    
    ' Set method and add body
    Request.Method = HttpPatch
    Set Request.Body = Values
    
    Set UpdateRequest = Request
End Function

''
' Setup generic query request
' --------------------------------------- '
Public Function QueryRequest(query As String) As WebRequest
    Dim Request As New WebRequest
    Request.Resource = "services/data/{ApiVersion}/query/"
    Request.AddUrlSegment "ApiVersion", ApiVersion
    Request.AddQuerystringParam "q", query
    
    Set QueryRequest = Request
End Function

 
 
