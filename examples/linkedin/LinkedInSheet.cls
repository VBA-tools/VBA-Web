VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkedInSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub LoadProfile()
    Me.[LinkedInFirstName] = "Loading..."
    Me.[LinkedInLastName] = ""
    Me.[LinkedInHeadline] = ""
    
    LinkedInSheet.LoadedProfile LinkedIn.GetProfile
End Sub

Public Sub LoadedProfile(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.Ok Then
        Me.[LinkedInFirstName] = Response.Data("firstName")
        Me.[LinkedInLastName] = Response.Data("lastName")
        Me.[LinkedInHeadline] = Response.Data("headline")
    Else
        Me.[LinkedInFirstName] = "Error: " & Response.Content
    End If
End Sub
