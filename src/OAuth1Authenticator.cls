VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuth1Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IAuthenticator
''
' OAuth1 Authenticator v0.2.2
' (c) Tim Hall - https://github.com/timhall/ExcelHelpers
'
' Utilize OAuth1 authentication
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @version 0.2.2
' @date 20121024
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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

Private Sub IAuthenticator_BeforeExecute(request As RestRequest)
	' Add authorization header to request
    Call request.AddHeader("Authorization", CreateHeader(request))
End Sub

Private Sub IAuthenticator_HttpOpen(http As MSXML2.IXMLHTTPRequest, request As RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Standard http open
    Call http.Open(request.MethodName(), request.FullUrl(BaseUrl), useAsync)
End Sub

Private Function CreateHeader(request As RestRequest) As String
    Dim header As String
    Dim nonce As String
    Dim timestamp As String
    Dim base As String
    Dim signingKey As String
    Dim signature As String
    
	' Create needed parts of authorization header
    nonce = CreateNonce()
    timestamp = CreateTimestamp()
    base = CreateBaseString(nonce, timestamp, request)
    signingKey = CreateSigningKey()
    signature = CreateSignature(base, signingKey)
    
    ' Add realm (if exists)
    If Me.Realm <> "" Then
        header = "OAuth realm=" & Chr(34) & Me.Realm & Chr(34) & ", "
    Else
        header = "OAuth "
    End If
    
    ' Construct header parts
    ' [OAuth Core 1.0 Revision A](http://oauth.net/core/1.0a/)
    header = header & "oauth_consumer_key=" & Chr(34) & Me.ConsumerKey & Chr(34)
    header = header & ", oauth_nonce=" & Chr(34) & nonce & Chr(34)
    header = header & ", oauth_signature=" & Chr(34) & URLEncode(signature) & Chr(34)
    header = header & ", oauth_signature_method=" & Chr(34) & SignatureMethod & Chr(34)
    header = header & ", oauth_timestamp=" & Chr(34) & timestamp & Chr(34)
    header = header & ", oauth_token=" & Chr(34) & Me.Token & Chr(34)
    header = header & ", oauth_version=" & Chr(34) & "1.0" & Chr(34)
    CreateHeader = header
    Debug.Print (vbNewLine & "OAuth1 Authorization: " & header)
End Function

Private Function CreateBaseString(nonce As String, timestamp As String, request As RestRequest) As String
    Dim base As String
    Dim paramKey As Variant
    
    base = "oauth_consumer_key" & "=" & Me.ConsumerKey
    base = base & "&" & "oauth_nonce" & "=" & nonce
    base = base & "&" & "oauth_signature_method" & "=" & SignatureMethod
    base = base & "&" & "oauth_timestamp" & "=" & timestamp
    base = base & "&" & "oauth_token" & "=" & Me.Token
    base = base & "&" & "oauth_version=1.0"
    If Not IsMissing(request.Parameters) And Not request.Parameters Is Nothing Then
        If request.Parameters.count > 0 Then
            For Each paramKey In request.Parameters.keys
                base = base & "&" & paramKey & "=" & URLEncode(request.Parameters(paramKey))
            Next paramKey
        End If
    End If
    
    CreateBaseString = request.MethodName() & "&" & URLEncode(request.FormattedResource()) & "&" & URLEncode(base)
    ' Debug.Print (vbNewLine & "Base string: " & CreateOAuth1BaseString & vbNewLine)
End Function

Private Function CreateSigningKey() As String
    CreateSigningKey = Me.ConsumerSecret & "&" & Me.TokenSecret
End Function

Private Function CreateSignature(base As String, signingKey As String) As String
    CreateSignature = Base64_HMACSHA1(base, signingKey)
    ' Debug.Print (vbNewLine & "Signature: " & CreateOAuth1Signature & vbNewLine)
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

