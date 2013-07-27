VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalyticsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub DownloadAnalytics()
    
    Me.[GAMessage] = "Loading"
    Me.[GAVisits] = ""
    Me.[GABounces] = ""
    
    Analytics.GAClient.ExecuteAsync Analytics.AnalyticsRequest(Me.[GAProfileId], Me.[GAStartDate], Me.[GAEndDate]), "AnalyticsSheet.OutputAnalytics"
    
End Sub

Public Sub OutputAnalytics(Response As RestResponse)

    If Response.StatusCode = Ok Then
        Me.[GAMessage] = "Loaded Asynchronously!"
        Me.[GAVisits] = Response.Data("totalsForAllResults")("ga:visits")
        Me.[GABounces] = Response.Data("totalsForAllResults")("ga:bounces")
    Else
        Me.[GAMessage] = "Error " & Response.StatusCode & ": " & Response.Content
    End If

End Sub


