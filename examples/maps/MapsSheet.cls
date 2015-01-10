VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub FindDrivingDistance()
    ' Loading
    Me.[Distance] = "Loading"
    Me.[Message] = ""

    ' Get directions for given origin and destination and process async
    OutputDrivingDistance Maps.MapsClient.Execute(Maps.DirectionsRequest(Me.[Origin], Me.[Destination]))
End Sub

Public Sub OutputDrivingDistance(Response As WebResponse)
    If Response.StatusCode = Ok Then
        Dim Route As Dictionary
        Dim Duration As String
        Dim Distance As String
        Dim StartAddress As String
        Dim EndAddress As String
        
        If Response.Data("status") <> "OK" Then
            Me.[Distance] = "No Results"
        Else
            Set Route = Response.Data("routes")(1)("legs")(1)
            Duration = Route("duration")("text")
            Distance = Route("distance")("text")
            
            Me.[Distance] = Distance & ", " & Duration
            ' Me.[Message] = "Found Asynchronously!"
        End If
    Else
        Me.[Message] = "Error: " & Response.Content
    End If
End Sub
