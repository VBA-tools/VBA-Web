VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwitterAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Twitter Authenticator v2.0.6
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for application-only authentication in Twitter's V1.1 REST API
'
' - https://dev.twitter.com/docs/auth/application-only-auth
' - https://github.com/timhall/Excel-REST/wiki/Implementing-your-own-IAuthenticator
'
' @implements: IAuthenticator v3.*
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IAuthenticator
Option Explicit

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
            RestHelpers.LogDebug "Successfully received token: " & RestHelpers.Obfuscate(pToken), "TwitterAuthenticator.Token"
        Else
            RestHelpers.LogError "Failed to load token: " & Response.StatusCode & " - " & Response.Content, "TwitterAuthenticator.Token"
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

' ============================================= '
' Private Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {RestClient} Client The client that is about to execute the request
' @param {RestRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IAuthenticator_BeforeExecute(ByVal Client As RestClient, ByRef Request As RestRequest)
    Request.AddHeader "Authorization", "Bearer " & Me.Token
End Sub

''
' Hook for taking action after request has been executed
'
' @param {RestClient} Client The client that executed request
' @param {RestRequest} Request The request that was just executed
' @param {RestResponse} Response to request
' --------------------------------------------- '

Private Sub IAuthenticator_AfterExecute(ByVal Client As RestClient, ByVal Request As RestRequest, ByRef Response As RestResponse)
    
End Sub

''
' Hook for overriding standard http open (used for HTTP Basic)
'
' @param {MSXML2.IXMLHTTPRequest} http
' @parma {RestClient} Client The client that is about to open request
' @param {RestRequest} Request The request about to be opened
' @param {String} BaseUrl
' @param {Boolean} [useAsync=False]
' --------------------------------------------- '

Private Sub IAuthenticator_HttpOpen(ByRef Http As Object, ByVal Client As RestClient, ByRef Request As RestRequest, BaseUrl As String, Optional UseAsync As Boolean = False)
    ' Use modified http open (if necessary)

    ' Perform standard http open
    Call Http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), UseAsync)
    RestHelpers.LogDebug "Http.Open " & Request.MethodName() & Request.FullUrl(BaseUrl), "FacebookAuthenticator.HttpOpen"
End Sub

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
