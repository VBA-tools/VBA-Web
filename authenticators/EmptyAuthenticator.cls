VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmptyAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IAuthenticator

Public Sub Setup()
    ' Define any user-specific variables needed for authentication
End Sub

Private Sub IAuthenticator_BeforeExecute(Request As RestRequest)
' (Used to add any required fields to the `Request` before it is executed)

    ' Leave blank to pass Request through unmodified
End Sub

Private Sub IAuthenticator_HttpOpen( _
    http As MSXML2.IXMLHTTPRequest, _
    Request As RestRequest, _
    BaseUrl As String, _
    Optional useAsync As Boolean = False)
' (Used to open the given http request, making any necessary modifications)

    ' Perform standard http open
    ' Call http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), useAsync)
End Sub
