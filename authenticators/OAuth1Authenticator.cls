VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuth1Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' OAuth1 Authenticator v2.0.2
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Utilize OAuth1 authentication
'
' @implements: IAuthenticator v3.*
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IAuthenticator
Option Explicit

' Declarations
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME)

' Constants and private variables
Private Const SignatureMethod As String = "HMAC-SHA1"

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private Enum TIME_ZONE
    TIME_ZONE_ID_INVALID = 0
    TIME_ZONE_STANDARD = 1
    TIME_ZONE_DAYLIGHT = 2
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ConsumerKey As String
Public ConsumerSecret As String
Public Token As String
Public TokenSecret As String
Public Realm As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Helper for setting required parameters at once
'
' @param {String} ConsumerKey
' @param {String} ConsumerSecret
' @param {String} Token
' @param {String} TokenSecret
' @param {String} [Realm]
' --------------------------------------------- '

Public Sub Setup(ConsumerKey As String, ConsumerSecret As String, _
    Token As String, TokenSecret As String, Optional Realm As String = "")
    
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
    Me.Token = Token
    Me.TokenSecret = TokenSecret
    Me.Realm = Realm
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
    ' Add authorization header to request
    Request.AddHeader "Authorization", CreateHeader(Request)
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
    ' Perform standard http open
    Call Http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), UseAsync)
End Sub

Private Function CreateHeader(Request As RestRequest) As String
    Dim Header As String
    Dim nonce As String
    Dim timestamp As String
    Dim base As String
    Dim signingKey As String
    Dim signature As String
    
    ' Create needed parts of authorization header
    nonce = CreateNonce()
    timestamp = CreateTimestamp()
    base = CreateBaseString(nonce, timestamp, Request)
    signingKey = CreateSigningKey()
    signature = CreateSignature(base, signingKey)
    
    ' Add realm (if exists)
    If Me.Realm <> "" Then
        Header = "OAuth realm=" & Chr(34) & Me.Realm & Chr(34) & ", "
    Else
        Header = "OAuth "
    End If
    
    ' Construct header parts
    ' [OAuth Core 1.0 Revision A](http://oauth.net/core/1.0a/)
    Header = Header & "oauth_consumer_key=" & Chr(34) & Me.ConsumerKey & Chr(34)
    Header = Header & ", oauth_nonce=" & Chr(34) & nonce & Chr(34)
    Header = Header & ", oauth_signature=" & Chr(34) & URLEncode(signature) & Chr(34)
    Header = Header & ", oauth_signature_method=" & Chr(34) & SignatureMethod & Chr(34)
    Header = Header & ", oauth_timestamp=" & Chr(34) & timestamp & Chr(34)
    Header = Header & ", oauth_token=" & Chr(34) & Me.Token & Chr(34)
    Header = Header & ", oauth_version=" & Chr(34) & "1.0" & Chr(34)
    CreateHeader = Header
End Function

Private Function CreateBaseString(nonce As String, timestamp As String, Request As RestRequest) As String
    Dim base As String
    Dim paramKey As Variant
    
    base = "oauth_consumer_key" & "=" & Me.ConsumerKey
    base = base & "&" & "oauth_nonce" & "=" & nonce
    base = base & "&" & "oauth_signature_method" & "=" & SignatureMethod
    base = base & "&" & "oauth_timestamp" & "=" & timestamp
    base = base & "&" & "oauth_token" & "=" & Me.Token
    base = base & "&" & "oauth_version=1.0"
    If Not IsMissing(Request.Parameters) And Not Request.Parameters Is Nothing Then
        If Request.Parameters.count > 0 Then
            For Each paramKey In Request.Parameters.keys
                base = base & "&" & paramKey & "=" & URLEncode(Request.Parameters(paramKey))
            Next paramKey
        End If
    End If
    
    CreateBaseString = Request.MethodName() & "&" & URLEncode(Request.FormattedResource()) & "&" & URLEncode(base)
End Function

Private Function CreateSigningKey() As String
    CreateSigningKey = Me.ConsumerSecret & "&" & Me.TokenSecret
End Function

Private Function CreateSignature(base As String, signingKey As String) As String
    CreateSignature = Base64_HMACSHA1(base, signingKey)
End Function

Private Function CreateTimestamp() As String
    CreateTimestamp = CStr(DateDiff("s", #1/1/1970#, GetGMT()))
End Function

' ======================================================================================== '
'
' Time Calculations
'
' ======================================================================================== '

Function GetGMT(Optional LocalTime As Date) As Date
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' http://www.cpearson.com/excel/TimeZoneAndDaylightTime.aspx
    ' ConvertLocalToGMT
    ' This function returns the GMT based on LocalTime, if provided.
    ' If LocalTime is not equal to 0, the GMT corresponding to LocalTime
    ' is returned. If LocalTime is 0, the GMT corresponding to the local
    ' time is returned. Since GMT isn't affected by DST, we need to
    ' subtract 1 hour if we are presently in DST.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim T As Date
    Dim TZI As TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    Dim GMT As Date
    
    If LocalTime <= 0 Then
        T = Now
    Else
        T = LocalTime
    End If
    DST = GetTimeZoneInformation(TZI)
    GMT = T + TimeSerial(0, TZI.Bias, 0) - IIf(DST = TIME_ZONE_DAYLIGHT, TimeSerial(1, 0, 0), 0)
    GetGMT = GMT
End Function
