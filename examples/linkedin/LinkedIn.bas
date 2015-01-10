Attribute VB_Name = "LinkedIn"
Private pLinkedInClient As WebClient
Private pLinkedInAPIKey As String
Private pLinkedInAPISecret As String
Private pLinkedInUserToken As String
Private pLinkedInUserSecret As String

Private Property Get LinkedInAPIKey() As String
    If pLinkedInAPIKey = "" Then
        If Credentials.Loaded Then
            pLinkedInAPIKey = Credentials.Values("LinkedIn")("api_key")
        Else
            pLinkedInAPIKey = InputBox("Please Enter LinkedIn API Key")
        End If
    End If
    
    LinkedInAPIKey = pLinkedInAPIKey
End Property

Private Property Get LinkedInAPISecret() As String
    If pLinkedInAPISecret = "" Then
        If Credentials.Loaded Then
            pLinkedInAPISecret = Credentials.Values("LinkedIn")("api_secret")
        Else
            pLinkedInAPISecret = InputBox("Please Enter LinkedIn API Secret")
        End If
    End If
    
    LinkedInAPISecret = pLinkedInAPISecret
End Property

Private Property Get LinkedInUserToken() As String
    If pLinkedInUserToken = "" Then
        If Credentials.Loaded Then
            pLinkedInUserToken = Credentials.Values("LinkedIn")("user_token")
        Else
            pLinkedInUserToken = InputBox("Please Enter LinkedIn User Token")
        End If
    End If
    
    LinkedInUserToken = pLinkedInUserToken
End Property

Private Property Get LinkedInUserSecret() As String
    If pLinkedInUserSecret = "" Then
        If Credentials.Loaded Then
            pLinkedInUserSecret = Credentials.Values("LinkedIn")("user_secret")
        Else
            pLinkedInUserSecret = InputBox("Please Enter LinkedIn User Secret")
        End If
    End If
    
    LinkedInUserSecret = pLinkedInUserSecret
End Property

Private Property Get LinkedInClient() As WebClient
    If pLinkedInClient Is Nothing Then
        Set pLinkedInClient = New WebClient
        pLinkedInClient.BaseUrl = "http://api.linkedin.com/v1/"
        
        Dim Auth As New OAuth1Authenticator
        Auth.Setup _
            ConsumerKey:=LinkedInAPIKey, _
            ConsumerSecret:=LinkedInAPISecret, _
            Token:=LinkedInUserToken, _
            TokenSecret:=LinkedInUserSecret
        Set pLinkedInClient.Authenticator = Auth
    End If
    
    Set LinkedInClient = pLinkedInClient
End Property

Public Function GetProfile(Optional Callback As String = "") As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "people/~?format=json"
    
'    If Callback <> "" Then
'        LinkedInClient.ExecuteAsync Request, Callback
'    Else
        Set GetProfile = LinkedInClient.Execute(Request)
'    End If
End Function
