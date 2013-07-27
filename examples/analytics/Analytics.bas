Attribute VB_Name = "Analytics"
Private pGAClient As RestClient
Private pGAClientId As String
Private pGAClientSecret As String

' Implement caching for Client Id, Client Secret, and RestClient

Private Property Get GAClientId() As String
    If pGAClientId = "" Then
        pGAClientId = InputBox("Please Enter Google API Client Id")
    End If
    
    GAClientId = pGAClientId
End Property
Private Property Get GAClientSecret() As String
    If pGAClientSecret = "" Then
        pGAClientSecret = InputBox("Please Enter Google API Client Secret")
    End If
    
    GAClientSecret = pGAClientSecret
End Property

Public Property Get GAClient() As RestClient
    If pGAClient Is Nothing Then
        Set pGAClient = New RestClient
        pGAClient.BaseUrl = "https://www.googleapis.com/analytics/v3"
        
        Dim Auth As New GoogleAuthenticator
        Set pGAClient.Authenticator = Auth
        Call Auth.Setup(GAClientId, GAClientSecret)
        Auth.Scope = Array("https://www.googleapis.com/auth/analytics.readonly")
        Call Auth.Login
    End If
    
    Set GAClient = pGAClient
End Property

Public Function AnalyticsRequest(ProfileId As String, StartDate As Date, EndDate As Date) As RestRequest
    
    Set AnalyticsRequest = New RestRequest
    AnalyticsRequest.Resource = "data/ga"
    AnalyticsRequest.Method = httpGET
    
    AnalyticsRequest.AddQuerystringParam "ids", "ga:" & ProfileId
    AnalyticsRequest.AddQuerystringParam "start-date", Format(StartDate, "yyyy-mm-dd")
    AnalyticsRequest.AddQuerystringParam "end-date", Format(EndDate, "yyyy-mm-dd")
    AnalyticsRequest.AddQuerystringParam "metrics", "ga:visits,ga:bounces"
    
End Function
