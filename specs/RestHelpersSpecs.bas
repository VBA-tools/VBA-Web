Attribute VB_Name = "RestHelpersSpecs"
''
' RestHelpersSpecs
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Specs for RestHelpers
'
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "RestHelpers"
    
    Dim json As String
    Dim Parsed As Object
    Dim Obj As Object
    Dim Coll As Collection
    Dim A As Object
    Dim B As Object
    Dim Combined As Object
    Dim Whitelist As Variant
    Dim Filtered As Object
    
    With Specs.It("should parse json")
        json = "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2]}"
        Set Parsed = RestHelpers.ParseJSON(json)
        
        .Expect(Parsed).ToBeDefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed("a")).ToEqual 1
            .Expect(Parsed("b")).ToEqual 3.14
            .Expect(Parsed("c")).ToEqual "Howdy!"
            .Expect(Parsed("d")).ToEqual True
            .Expect(Parsed("e").count).ToEqual 2
        End If
        
        json = "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""}]"
        Set Parsed = RestHelpers.ParseJSON(json)
        
        .Expect(Parsed).ToBeDefined
        If Not Parsed Is Nothing Then
            .Expect(Parsed(1)).ToEqual 1
            .Expect(Parsed(2)).ToEqual 3.14
            .Expect(Parsed(3)).ToEqual "Howdy!"
            .Expect(Parsed(4)).ToEqual True
            .Expect(Parsed(5).count).ToEqual 2
            .Expect(Parsed(6)("a")).ToEqual "Howdy!"
        End If
    End With
    
    With Specs.It("should convert to json")
        Set Obj = CreateObject("Scripting.Dictionary")
        Obj.Add "a", 1
        Obj.Add "b", 3.14
        Obj.Add "c", "Howdy!"
        Obj.Add "d", True
        Obj.Add "e", Array(1, 2)
        
        json = RestHelpers.ConvertToJSON(Obj)
        .Expect(json).ToEqual "{""a"":1,""b"":3.14,""c"":""Howdy!"",""d"":true,""e"":[1,2]}"
        
        Set Obj = CreateObject("Scripting.Dictionary")
        Obj.Add "a", "Howdy!"
        
        Set Coll = New Collection
        Coll.Add 1
        Coll.Add 3.14
        Coll.Add "Howdy!"
        Coll.Add True
        Coll.Add Array(1, 2)
        Coll.Add Obj
        
        json = RestHelpers.ConvertToJSON(Coll)
        .Expect(json).ToEqual "[1,3.14,""Howdy!"",true,[1,2],{""a"":""Howdy!""}]"
    End With
    
    With Specs.It("should url encode values")
        .Expect(RestHelpers.URLEncode(" !""#$%&'")).ToEqual "%20%21%22%23%24%25%26%27"
    End With
    
    With Specs.It("should join url with /")
        .Expect(RestHelpers.JoinUrl("a", "b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a/", "b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a", "/b")).ToEqual "a/b"
        .Expect(RestHelpers.JoinUrl("a/", "/b")).ToEqual "a/b"
    End With
    
    With Specs.It("should combine objects, with overwrite option")
        Set A = CreateObject("Scripting.Dictionary")
        Set B = CreateObject("Scripting.Dictionary")
        
        A.Add "a", 1
        A.Add "b", 3.14
        B.Add "b", 4.14
        B.Add "c", "Howdy!"
        
        Set Combined = RestHelpers.CombineObjects(A, B)
        .Expect(Combined("a")).ToEqual 1
        .Expect(Combined("b")).ToEqual 4.14
        .Expect(Combined("c")).ToEqual "Howdy!"
        
        Set Combined = RestHelpers.CombineObjects(A, B, OverwriteOriginal:=False)
        .Expect(Combined("a")).ToEqual 1
        .Expect(Combined("b")).ToEqual 3.14
        .Expect(Combined("c")).ToEqual "Howdy!"
    End With
    
    With Specs.It("should filter object by whitelist")
        Set Obj = CreateObject("Scripting.Dictionary")
        Obj.Add "a", 1
        Obj.Add "b", 3.14
        Obj.Add "dangerous", "Howdy!"
        
        Whitelist = Array("a", "b")
        
        Set Filtered = RestHelpers.FilterObject(Obj, Whitelist)
        .Expect(Obj.Exists("dangerous")).ToEqual True
        .Expect(Filtered.Exists("a")).ToEqual True
        .Expect(Filtered.Exists("b")).ToEqual True
        .Expect(Filtered.Exists("dangerous")).ToEqual False
    End With
    
    With Specs.It("should encode string to base64")
        .Expect(RestHelpers.EncodeStringToBase64("Howdy!")).ToEqual "SG93ZHkh"
    End With
    
    With Specs.It("should create Nonce of specified length")
        .Expect(Len(RestHelpers.CreateNonce)).ToEqual 32
        .Expect(Len(RestHelpers.CreateNonce(20))).ToEqual 20
    End With
    
    InlineRunner.RunSuite Specs
End Function
