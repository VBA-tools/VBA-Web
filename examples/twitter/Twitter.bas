Attribute VB_Name = "Twitter"
Private pTwitterClient As WebClient
Private pTwitterKey As String
Private pTwitterSecret As String

' Implement caching for Consumer Key, Consumer Secret, and WebClient

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

Private Property Get TwitterClient() As WebClient
    If pTwitterClient Is Nothing Then

        Set pTwitterClient = New WebClient
        pTwitterClient.BaseUrl = "https://api.twitter.com/1.1/"
        
        Dim Auth As New TwitterAuthenticator
        Auth.Setup _
            ConsumerKey:=TwitterKey, _
            ConsumerSecret:=TwitterSecret
        Set pTwitterClient.Authenticator = Auth
    End If
    
    Set TwitterClient = pTwitterClient
End Property



Private Function SearchTweetsRequest(query As String) As WebRequest
    Set SearchTweetsRequest = New WebRequest
    SearchTweetsRequest.Resource = "search/tweets.{format}"
    
    SearchTweetsRequest.Format = Json
    SearchTweetsRequest.AddUrlSegment "format", "json"
    SearchTweetsRequest.AddQuerystringParam "q", query
    SearchTweetsRequest.AddQuerystringParam "lang", "en"
    SearchTweetsRequest.AddQuerystringParam "count", 20
    SearchTweetsRequest.Method = HttpGet
End Function

Public Function SearchTwitter(query As String) As WebResponse
    Set SearchTwitter = TwitterClient.Execute(SearchTweetsRequest(query))
End Function

'Public Sub SearchTwitterAsync(query As String, Callback As String)
'    TwitterClient.ExecuteAsync SearchTweetsRequest(query), Callback
'End Sub
    
