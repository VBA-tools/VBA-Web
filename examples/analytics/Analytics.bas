Attribute VB_Name = "Analytics"
Private pGAClient As WebClient
Private pGAClientId As String
Private pGAClientSecret As String

' Implement caching for Client Id, Client Secret, and WebClient

Private Property Get GAClientId() As String
    If pGAClientId = "" Then
        If Credentials.Loaded Then
            pGAClientId = Credentials.Values("Google")("id")
        Else
            pGAClientId = InputBox("Please Enter Google API Client Id")
        End If
    End If
    
    GAClientId = pGAClientId
End Property
Private Property Get GAClientSecret() As String
    If pGAClientSecret = "" Then
        If Credentials.Loaded Then
            pGAClientSecret = Credentials.Values("Google")("secret")
        Else
            pGAClientSecret = InputBox("Please Enter Google API Client Secret")
        End If
    End If
    
    GAClientSecret = pGAClientSecret
End Property

Public Property Get GAClient() As WebClient
    If pGAClient Is Nothing Then
        Set pGAClient = New WebClient
        pGAClient.BaseUrl = "https://www.googleapis.com/analytics/v3"
        
        Dim Auth As New GoogleAuthenticator
        Auth.Setup GAClientId, GAClientSecret
        Auth.AddScope "analytics.readonly"
        Call Auth.Login
        
        Set pGAClient.Authenticator = Auth
    End If
    
    Set GAClient = pGAClient
End Property

Public Function AnalyticsRequest(ProfileId As String, StartDate As Date, EndDate As Date) As WebRequest
    
    If ProfileId = "" And Credentials.Loaded Then
        ProfileId = Credentials.Values("Google")("profile")
    End If
    
    Set AnalyticsRequest = New WebRequest
    AnalyticsRequest.Resource = "data/ga"
    AnalyticsRequest.Method = WebMethod.HttpGet
    
    AnalyticsRequest.AddQuerystringParam "ids", "ga:" & ProfileId
    AnalyticsRequest.AddQuerystringParam "start-date", Format(StartDate, "yyyy-mm-dd")
    AnalyticsRequest.AddQuerystringParam "end-date", Format(EndDate, "yyyy-mm-dd")
    AnalyticsRequest.AddQuerystringParam "metrics", "ga:visits,ga:bounces"
    
End Function
