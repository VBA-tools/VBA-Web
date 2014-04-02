Attribute VB_Name = "Twitter"
Private pTwitterClient As RestClient
Private pTwitterKey As String
Private pTwitterSecret As String

' Implement caching for Consumer Key, Consumer Secret, and RestClient

Private Property Get TwitterKey() As String
    If pTwitterKey = "" Then
        If Credentials.Loaded Then
            pTwitterKey = Credentials.Values("Twitter")("key")
        Else
            pTwitterKey = InputBox("Please Enter Twitter Consumer Key")
        End If
    End If
    
    TwitterKey = pTwitterKey
End Property
Private Property Get TwitterSecret() As String
    If pTwitterSecret = "" Then
        If Credentials.Loaded Then
            pTwitterSecret = Credentials.Values("Twitter")("secret")
        Else
            pTwitterSecret = InputBox("Please Enter Twitter Consumer Secret")
        End If
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



Private Function SearchTweetsRequest(query As String) As RestRequest
    Set SearchTweetsRequest = New RestRequest
    SearchTweetsRequest.Resource = "search/tweets.{format}"
    
    SearchTweetsRequest.Format = json
    SearchTweetsRequest.AddParameter "q", query
    SearchTweetsRequest.AddParameter "lang", "en"
    SearchTweetsRequest.AddParameter "count", 20
    SearchTweetsRequest.Method = httpGET
End Function

Public Function SearchTwitter(query As String) As RestResponse
    Set SearchTwitter = TwitterClient.Execute(SearchTweetsRequest(query))
End Function

Public Sub SearchTwitterAsync(query As String, Callback As String)
    TwitterClient.ExecuteAsync SearchTweetsRequest(query), Callback
End Sub
    
