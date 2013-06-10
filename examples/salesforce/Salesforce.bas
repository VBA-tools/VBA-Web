Attribute VB_Name = "Salesforce"
' Salesforce client with RESTful methods
' http://www.salesforce.com/us/developer/docs/api_rest/index.htm

Private Const ApiVersion As String = "v26.0"
Private pClient As RestClient

Public ConsumerKey As String
Public ConsumerSecret As String
Public Username As String
Public Password As String
Public SecurityToken As String

Public Property Get Client() As RestClient
    If pClient Is Nothing Then
        Set pClient = New RestClient
        pClient.BaseUrl = "https://na1.salesforce.com/"
        
        ' Setup OAuth2 authentication
        Dim Auth As New OAuth2Authenticator
        Auth.Setup _
            ClientId:=ConsumerKey, _
            ClientSecret:=ConsumerSecret, _
            Username:=Username, _
            Password:=Password & SecurityToken
        
        Auth.SetupTokenUrl "https://login.salesforce.com/services/oauth2/token?grant_type=password"
        Set pClient.Authenticator = Auth
    End If
    
    Set Client = pClient
End Property

''
' Get generic object (async and sync)
' --------------------------------------- '
Public Sub GetObjectAsync(ObjectName As String, ObjectId As String, _
    Callback As String, ParamArray CallbackArgs() As Variant)
    
    Call Client.ExecuteAsync(ObjectRequest(ObjectName, ObjectId), Callback, CallbackArgs)
End Sub

Public Function GetObject(ObjectName As String, ObjectId As String) As RestResponse
    Set GetObject = Client.Execute(ObjectRequest(ObjectName, ObjectId))
End Function

''
' Update generic object (async and sync)
' --------------------------------------- '
Public Sub UpdateObjectAsync(ObjectName As String, ObjectId As String, values As Dictionary, _
    Callback As String, ParamArray CallbackArgs() As Variant)
    
    Call Client.ExecuteAsync(UpdateRequest(ObjectName, ObjectId, values), Callback, CallbackArgs)
End Sub

Public Function UpdateObject(ObjectName As String, ObjectId As String, values As Dictionary) As RestResponse
    Set UpdateObject = Client.Execute(UpdateRequest(ObjectName, ObjectId, values))
End Function

''
' Run query (async and sync)
' --------------------------------------- '
Public Sub RunQueryAsync(query As String, _
    Callback As String, ParamArray CallbackArgs() As Variant)
    
    Call Client.ExecuteAsync(QueryRequest(query), Callback, CallbackArgs)
End Sub

Public Function RunQuery(query As String) As RestResponse
    Set RunQuery = Client.Execute(QueryRequest(query))
End Function

''
' Get overview info for the specified object
' --------------------------------------- '
Public Function GetObjectInfo(ObjectName As String) As RestResponse
    Dim Request As RestRequest
    Set Request = ObjectRequest(ObjectName, "describe")
    
    Set GetObjectInfo = Client.Execute(Request)
End Function

''
' Setup generic object request
' --------------------------------------- '
Public Function ObjectRequest(ObjectName As String, ObjectId As String) As RestRequest
    Dim Request As New RestRequest
    Request.Resource = "services/data/{ApiVersion}/sobjects/{ObjectName}/{ObjectId}"
    Request.AddUrlSegment "ApiVersion", ApiVersion
    Request.AddUrlSegment "ObjectName", ObjectName
    Request.AddUrlSegment "ObjectId", ObjectId
    Set ObjectRequest = Request
End Function

''
' Setup generic update request
' --------------------------------------- '
Public Function UpdateRequest(ObjectName As String, ObjectId As String, values As Dictionary) As RestRequest
    Dim Request As RestRequest
    Set Request = ObjectRequest(ObjectName, ObjectId)
    
    ' Remove Id from values for update
    If values.Exists("Id") Then values.Remove "Id"
    
    ' Set method and add body
    Request.Method = httpPATCH
    Request.AddBody values
    
    Set UpdateRequest = Request
End Function

''
' Setup generic query request
' --------------------------------------- '
Public Function QueryRequest(query As String) As RestRequest
    Dim Request As New RestRequest
    Request.Resource = "services/data/{ApiVersion}/query/"
    Request.AddUrlSegment "ApiVersion", ApiVersion
    Request.AddParameter "q", query
    
    Set QueryRequest = Request
End Function

 
 
