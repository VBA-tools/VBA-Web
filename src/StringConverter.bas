Attribute VB_Name = "StringConverter"
''
' Module used for converting string to UTF8 byte array.
'
' Code taken from page http://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html
' Modified to work in 64-bit Excel using guide at http://stackoverflow.com/questions/21982682/code-does-not-work-on-64-bit-office
'
' Used in WebHelpers.Base64Encode so that the HttpBasicAuthenticator works correctly with Scandinavian letters in username or password.
''

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As LongPtr, _
    ByVal dwFlags As LongPtr, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As LongPtr, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As LongPtr, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As LongPtr
    
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Variant
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
    
End Function

