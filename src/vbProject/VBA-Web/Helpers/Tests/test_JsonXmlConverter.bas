Attribute VB_Name = "test_JsonXmlConverter"
''
' VBA-Git Annotations
' https://github.com/VBA-Tools-v2/VBA-Git | https://radiuscore.co.nz
'
' @developmentmodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@testmodule
'@folder VBA-Web.Helpers.Tests
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Private Module

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TTest
    Assert As Object
    Fakes As Object
End Type

Private This As TTest

' ============================================= '
' Test Methods
' ============================================= '

'@TestMethod("WebHelpers")
Private Sub ConvertJsonToXml_SimpleJson()
    
    ' Arrange:
    
    ' Act:
    'WebHelpers.ConvertJsonToXml
    
    ' Assert:
    
End Sub

'@TestMethod("WebHelpers")
Private Sub ConvertXmlToJson_SimpleXml()

End Sub

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set This.Assert = CreateObject("Rubberduck.AssertClass")
    Set This.Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set This.Assert = Nothing
    Set This.Fakes = Nothing
End Sub
