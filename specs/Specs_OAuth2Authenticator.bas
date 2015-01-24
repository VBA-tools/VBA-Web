Attribute VB_Name = "Specs_OAuth2Authenticator"
''
' Specs_OAuth2Authenticator
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Specs for OAuth2Authenctiator
'
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "OAuth2Authenticator"
        
    Dim Client As New WebClient
    Dim Request As New WebRequest
    Dim Auth As New OAuth2Authenticator
    
    Dim ClientId As String
    Dim ClientSecret As String
    Dim Username As String
    Dim Password As String
    
    ' TODO
    
    InlineRunner.RunSuite Specs
End Function

