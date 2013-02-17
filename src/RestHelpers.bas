Attribute VB_Name = "RestHelpers"
''
' RestHelpers v0.2.2
' (c) Tim Hall - https://github.com/timhall/ExcelHelpers
'
' Common helpers RestClient
'
' @dependencies
'   JSONLib (http://code.google.com/p/vba-json/)
'   Microsoft XML, v3+
' @author tim.hall.engr@gmail.com
' @version 0.2.2
' @date 20121024
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Enum StatusCodes
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
    UnsupportedMediaType = 415
    InternalServerError = 500
    BadGateway = 502
    ServiceUnavailable = 503
    GatewayTimeout = 504
End Enum


' ============================================= '
' Shared Helpers
' ============================================= '

''
' Parse given JSON string into object (Dictionary or Collection)
'
' @param {String} jsonStr
' @return {Object} 
' --------------------------------------------- '

Public Function ParseJSON(jsonStr As String) As Object
    Dim lib As New JSONLib
    Set ParseJSON = lib.parse(jsonStr)
    Set lib = Nothing
End Function

''
' Convert object to JSON string
'
' @param {Object}
' @return {String}
' --------------------------------------------- '

Public Function ConvertToJSON(obj As Object) As String
    Dim lib As New JSONLib
    ConvertToJSON = lib.ToString(obj)
    Set lib = Nothing
End Function

''
' URL Encode the given raw values
'
' @param {Variant} rawVal The raw string to encode
' @param {Boolean} [spaceAsPlus=False] Use plus sign for encoded spaces (otherwise %20)
' @return {String} Encoded string
' --------------------------------------------- '

Public Function URLEncode(rawVal As Variant, Optional spaceAsPlus As Boolean = False) As String
    Dim urlVal As String
    Dim stringLen As Long
    
    urlVal = CStr(rawVal)
    stringLen = Len(urlVal)
    
    If stringLen > 0 Then
        ReDim result(stringLen) As String
        Dim i As Long, charCode As Integer
        Dim char As String, space As String
        
        ' Set space value
        If spaceAsPlus Then
            space = "+"
        Else
            space = "%20"
        End If
        
        ' Loop through string characters
        For i = 1 To stringLen
            ' Get character and ascii code
            char = Mid$(urlVal, i, 1)
            charCode = asc(char)
            Select Case charCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    ' Use original for AZaz09-._~
                    result(i) = char
                Case 32
                    ' Add space
                    result(i) = space
                Case 0 To 15
                    ' Convert to hex w/ leading 0
                    result(i) = "%0" & Hex(charCode)
                Case Else
                    ' Convert to hex
                    result(i) = "%" & Hex(charCode)
            End Select
        Next i
        URLEncode = Join(result, "")
    End If
End Function

''
' Combine two objects
'
' @param {Dictionary} origObj Original object to add values to
' @param {Dictionary} newObj New object containing values to add to original object
' @param {Boolean} [overwriteOriginal=True] Overwrite any values that already exist in the original object
' --------------------------------------------- '

Public Function CombineObjects(ByVal origObj As Dictionary, ByVal newObj As Dictionary, _
    Optional overwriteOriginal As Boolean = True) As Dictionary
    
    Dim combined As Dictionary
    Dim newKey As Variant
    
    If Not origObj Is Nothing Then
        Set combined = origObj
    Else
        Set combined = New Dictionary
    End If
    For Each newKey In newObj.keys()
        If combined.exists(newKey) And overwriteOriginal Then
            combined(newKey) = newObj(newKey)
        Else
            combined.Add newKey, newObj(newKey)
        End If
    Next newKey

    Set CombineObjects = combined
End Function

''
' Apply whitelist to given model to filter out unwanted key/values
'
' @param {Dictionary} original Original model to filter
' @param {Variant} whitelist Array of values to retain in the model
' --------------------------------------------- '

Public Function UpdateModel(ByVal original As Dictionary, whitelist As Variant) As Dictionary
    Dim updated As New Dictionary
    Dim i As Integer
    
    If IsArray(whitelist) Then
        For i = LBound(whitelist) To UBound(whitelist)
            If original.exists(whitelist(i)) Then
                updated.Add whitelist(i), original(whitelist(i))
            End If
        Next i
    ElseIf VarType(whitelist) = vbString Then
        If original.exists(whitelist) Then
            updated.Add whitelist, original(whitelist)
        End If
    End If
    
    Set UpdateModel = updated
End Function

' ======================================================================================== '
'
' Crytography and encoding
'
' ======================================================================================== '

'' 
' Generate a keyed hash value using the HMAC method and SHA1 algorithm
' [Does VBA have a Hash_HMAC](http://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac)
' 
' @param {String} sTextToHash
' @param {String} sSharedSecretKey 
' @return {String} 
' --------------------------------------------- '

Public Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey  As String) As String
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey )
    enc.key = SharedSecretKey

    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = EncodeBase64(bytes)
    Set asc = Nothing
    Set enc = Nothing
End Function

''
' Base64 encode data
'
' @param {Byte Array} arrData
' @return {String} Encoded string
' --------------------------------------------- '

Public Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML = New MSXML2.DOMDocument

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function

''
' Base64 encode string value
'
' @param {String} Data
' @return {String} Encoded string
' --------------------------------------------- '

Public Function EncodeStringToBase64(ByVal Data As String) As String
    Dim asc As Object
    Dim bytes() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    bytes = asc.Getbytes_4(Data)
    EncodeStringToBase64 = EncodeBase64(bytes)
    Set asc = Nothing
End Function

''
' Create random alphanumeric nonce
'
' @param {Integer} [NonceLength=32]
' @return {String} Randomly generated nonce
' --------------------------------------------- '

Public Function CreateNonce(Optional NonceLength As Integer = 32) As String
    Dim str As String
    Dim count As Integer
    Dim result As String
    Dim random As Integer
    
    str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    result = ""
    
    For count = 1 To NonceLength
        random = Int(((Len(str) - 1) * Rnd) + 1)
        result = result + Mid$(str, random, 1)
    Next
    CreateNonce = result
End Function
