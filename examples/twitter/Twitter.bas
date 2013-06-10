Attribute VB_Name = "Twitter"
Private pTwitterClient As RestClient
Private pTwitterKey As String
Private pTwitterSecret As String

' Implement caching for Consumer Key, Consumer Secret, and RestClient

Private Property Get TwitterKey() As String
    If pTwitterKey = "" Then
        pTwitterKey = InputBox("Please Enter Twitter Consumer Key")
    End If
    
    TwitterKey = pTwitterKey
End Property
Private Property Get TwitterSecret() As String
    If pTwitterSecret = "" Then
        pTwitterSecret = InputBox("Please Enter Twitter Consumer Secret")
    End If
    
    TwitterSecret = pTwitterSecret
End Property

Private Property Get TwitterClient() As RestClient
    If pTwitterClient Is Nothing Then

        Set pTwitterClient = New RestClient
        pTwitterClient.BaseUrl = "https://api.twitter.com/1.1/"
        
        Dim Auth As New TwitterAuthenticator
        Auth.Setup _
            ConsumerKey:=TwitterKey, _
            ConsumerSecret:=TwitterSecret
        Set pTwitterClient.Authenticator = Auth
    End If
    
    Set TwitterClient = pTwitterClient
End Property



Private Function SearchTweetsRequest(Query As String) As RestRequest
    Set SearchTweetsRequest = New RestRequest
    SearchTweetsRequest.Resource = "search/tweets.{format}"
    
    SearchTweetsRequest.Format = json
    SearchTweetsRequest.AddParameter "q", Query
    SearchTweetsRequest.AddParameter "lang", "en"
    SearchTweetsRequest.AddParameter "count", 20
    SearchTweetsRequest.Method = httpGET
End Function

Public Function SearchTwitter(Query As String) As RestResponse
    Set SearchTwitter = TwitterClient.Execute(SearchTweetsRequest(Query))
End Function

Public Sub SearchTwitterAsync(Query As String, Callback As String)
    TwitterClient.ExecuteAsync SearchTweetsRequest(Query), Callback
End Sub
    
