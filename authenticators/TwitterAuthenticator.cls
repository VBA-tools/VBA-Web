VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwitterAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IAuthenticator
''
' Twitter Authenticator v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for application-only authentication in Twitter's V1.1 REST API
'
' - https://dev.twitter.com/docs/auth/application-only-auth
' - https://github.com/timhall/Excel-REST/wiki/Implementing-your-own-IAuthenticator
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private pToken As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ConsumerKey As String
Public ConsumerSecret As String
Public CacheToken As Boolean

Public Property Get Token() As String
    On Error GoTo CleanUp

    If pToken = "" Or Not CacheToken Then
        Dim Client As New RestClient
        Dim Response As RestResponse
        
        ' Request a new token
        Client.BaseUrl = "https://api.twitter.com/"
        Set Response = Client.Execute(TokenRequest)
        
        ' Store token if successful, otherwise throw error
        If Response.StatusCode = Ok Then
            pToken = Response.Data("access_token")
        Else
            Err.Raise vbObjectError + Response.StatusCode, _
                Description:="Failed to load Bearer Token: " & Response.StatusCode & " - " & Response.Content
        End If
    End If
    
    ' Return token
    Token = pToken
    
CleanUp:

    If Not Client Is Nothing Then Set Client = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then Err.Raise Err.Number, Description:=Err.Description
    
End Property
Public Property Let Token(Value As String)
    pToken = Value
End Property

' ============================================= '
' Public Methods
' ============================================= '

Public Sub Setup(ConsumerKey As String, ConsumerSecret As String)
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
End Sub

Private Sub IAuthenticator_BeforeExecute(Request As RestRequest)
    Request.AddHeader "Authorization", "Bearer " & Me.Token
End Sub

Private Sub IAuthenticator_HttpOpen( _
    http As MSXML2.IXMLHTTPRequest, _
    Request As RestRequest, _
    BaseUrl As String, _
    Optional useAsync As Boolean = False)
    
    ' Perform standard http open
    Call http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), useAsync)
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Function TokenRequest() As RestRequest
    Set TokenRequest = New RestRequest
    TokenRequest.Resource = "oauth2/token"
    TokenRequest.Method = httpPOST
    
    ' Based on documentation:
    ' 1. Pass in Base64 encoded ConsumerKey and ConsumerSecret
    TokenRequest.AddHeader "Authorization", _
        "Basic " & RestHelpers.EncodeStringToBase64(Me.ConsumerKey & ":" & Me.ConsumerSecret)

    ' 2. Include grant_type in body
    TokenRequest.AddBodyString "grant_type=client_credentials"

    ' 3. Set content-type as form-urlencoded
    TokenRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
End Function

Private Sub Class_Initialize()
    ' Cache token by default
    Me.CacheToken = True
End Sub
