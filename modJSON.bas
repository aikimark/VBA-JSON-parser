Attribute VB_Name = "modJSON"
Option Explicit

Enum MachineState
    StartOfList = 1
    StartOfObject = 2
    EndOfObject = 3
    EndOfList = 4
End Enum

Function AddCurrentObjectToTopOfStackObject(parmObject As Object, parmStack As Collection) As Object
    Dim vTOSobject As Object
    Dim vKeys As Variant
    Set vTOSobject = parmStack(1)
    Select Case TypeName(vTOSobject)
        Case "Dictionary"
            vKeys = vTOSobject.keys
            Set vTOSobject(vKeys(UBound(vKeys))) = parmObject
        Case "Collection"
            vTOSobject.Add parmObject
    End Select
    Set AddCurrentObjectToTopOfStackObject = vTOSobject
    parmStack.Remove 1
End Function

Public Function parseJSON(parmJSON As String) As Object

    Dim oRE As Object
    Dim oRE_object As Object
    Dim oRE_list As Object
    Dim oMatches As Object
    Dim oM As Object
    Dim oSM As Object
    Dim lngM As Long
    Dim lngSM As Long
    Dim vItem As Variant
    
    Dim colMatchFirstIndexes As New Collection
    
    Dim strText As String
    
    Dim colStack As New Collection
    Dim oCurrentObject As Object
    Dim strToStype As String
    
    Dim strSectionText As String
    Const cStateChars As String = "[{}]"
    Dim vStateNames As Variant
    vStateNames = Array("StartOfList", "StartOfObject", "EndOfObject", "EndOfList")
    Dim lngDepth As Long
    
    strText = Trim(parmJSON)
    
    Set oRE_list = CreateObject("vbscript.regexp")
    oRE_list.Global = True
    oRE_list.Pattern = "(""((?:.|\"")+?)"")|([^ ,]+?)(?:,| |$)"
    Set oRE_object = CreateObject("vbscript.regexp")
    oRE_object.Global = True
    oRE_object.Pattern = """([^""]+)"": ?((""((?:\\""|.)+?)"")|([^ ,]*?))(?:,| |$)"
    Set oRE = CreateObject("vbscript.regexp")
    oRE.Global = True
    
    'preprocess the matches to ignore the escaped major delimiters
    oRE.Pattern = "(\\\[|\\{|\\}|\\\])" ' "(\[|{|}|\])"      '"(\[|{|: |}|\])"
    If oRE.test(strText) Then
        oRE.Pattern = "(\\\[|\\{|\\}|\\\]|\[|{|}|\])" ' "(\[|{|}|\])"      '"(\[|{|: |}|\])"
        Set oMatches = oRE.Execute(strText)
        For Each oM In oMatches
            If Len(oM) = 1 Then
                colMatchFirstIndexes.Add oM.firstindex + 1
            Else
                'Debug.Print oM.firstindex + 1, oM, "Skipped major delimiter"
            End If
        Next
    Else
        oRE.Pattern = "(\[|{|}|\])"       '"(\[|{|: |}|\])"
        Set oMatches = oRE.Execute(strText)
        For Each oM In oMatches
            colMatchFirstIndexes.Add oM.firstindex + 1
        Next
    End If
    Select Case True
        Case colMatchFirstIndexes.Count = 0
            MsgBox "no major delimiters, [|{|}|] , found", vbCritical
            Exit Function
        Case (colMatchFirstIndexes.Count Mod 2) <> 0
            MsgBox "odd number of major delimiters, [|{|}|] , found", vbCritical
            Exit Function
    End Select
    oRE.Pattern = "([\r\n])"
    For lngM = 1 To colMatchFirstIndexes.Count - 1
        Select Case InStr(cStateChars, Mid(strText, colMatchFirstIndexes(lngM), 1))
            Case 1  '[ -- start of list
                lngDepth = lngDepth + 1
                If oCurrentObject Is Nothing Then
                    Set oCurrentObject = New Collection
                    'Debug.Print "current object type: " & TypeName(oCurrentObject), "Stack size: " & colStack.Count
                Else
                    'push dicCurrentObject onto the stack
                    PushObjectOntoStack oCurrentObject, colStack
                    'create a new collection object
                    Set oCurrentObject = New Collection
                End If
            
            Case 2  '{ -- start of object
                lngDepth = lngDepth + 1
                If oCurrentObject Is Nothing Then
                    Set oCurrentObject = CreateObject("scripting.dictionary")
                    'Debug.Print "current object type: " & TypeName(oCurrentObject), "Stack size: " & colStack.Count
                Else
                    'push dicCurrentObject onto the stack
                    PushObjectOntoStack oCurrentObject, colStack
                    'create a new dictionary Object
                    Set oCurrentObject = CreateObject("scripting.dictionary")
                End If
            
            Case 3, 4   '} or ] -- end of object or list
                lngDepth = lngDepth - 1
                'add current object to top of stack object
                Set oCurrentObject = AddCurrentObjectToTopOfStackObject(oCurrentObject, colStack)

        End Select
        
        strSectionText = Mid(strText, colMatchFirstIndexes(lngM), colMatchFirstIndexes(lngM + 1) - colMatchFirstIndexes(lngM))
        strSectionText = Trim(Mid(oRE.Replace(strSectionText, ""), 2)) & " "
        Select Case TypeName(oCurrentObject)
            Case "Dictionary"
                Set oMatches = oRE_object.Execute(strSectionText)
            Case "Collection"
                Set oMatches = oRE_list.Execute(strSectionText)
        End Select
        
'        If colStack.Count = 0 Then
'            strToStype = "n/a"
'        Else
'            strToStype = TypeName(colStack(1))
'        End If
'        Debug.Print "current object type: " & TypeName(oCurrentObject), "Stack size: " & colStack.Count, "ToS object type: " & strToStype, "omatches.count: " & oMatches.Count, strSectionText
        If oMatches.Count <> 0 Then
            PopulateCurrentObject oMatches, oCurrentObject
        End If
    Next
    Set parseJSON = oCurrentObject
End Function

Public Function parseJSONfile(parmFilename As String) As Object
    Dim strText As String
    Dim oThing As Object
    Dim oFS, oTS
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

    Set oFS = CreateObject("scripting.filesystemobject")
    Set oTS = oFS.OpenTextFile(parmFilename, ForReading, True, TristateFalse)
    strText = oTS.ReadAll()
    oTS.Close
    
    Set oThing = parseJSON(Trim(strText))
    Set parseJSONfile = oThing
End Function

Function PopObjectFromStack(parmStack As Collection) As Object
    If parmStack.Count = 0 Then
        Set PopObjectFromStack = Nothing
    Else
        Set PopObjectFromStack = parmStack(1)
        parmStack.Remove 1
    End If
End Function

Sub PopulateCurrentObject(parmMatches As Object, parmCurObj As Object)
    Dim oM As Object
    Dim vValue As Variant
    Static oRE As Object
    Dim oUMatches As Object
    Dim oUM As Object
    Dim lngObjectType As Long
    Dim lngZoffset As Long
    Dim dtZoffset As Long
    
    If oRE Is Nothing Then
        Set oRE = CreateObject("vbscript.regexp")
        oRE.Global = True
        oRE.ignorecase = True
        oRE.Pattern = "\\U([0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f])"
    End If
    Select Case TypeName(parmCurObj)
        Case "Dictionary"
            lngObjectType = 3
        Case "Collection"
            lngObjectType = 1
    End Select
    For Each oM In parmMatches

'Print "0: " & oM.submatches(0), "1: " & oM.submatches(1), "2: " & oM.submatches(2)
'0: "XX0XXX"   1: XX0XXX     2:
'ToDo: The following condition misses lists of string values
'Consider this list parsing pattern: ("(?:(?:.|\\")+?)")|([^ ,]+?)(?:,| |$)
        If Left(oM.submatches(1), 1) = Chr(34) Then     'working with a string value
'will probably affect the following statement, too
            vValue = oM.submatches(lngObjectType)
            If InStr(vValue, "\") <> 0 Then
                vValue = Replace(vValue, "\""", """")
                vValue = Replace(vValue, "\\", "\")
                vValue = Replace(vValue, "\/", "/")
                vValue = Replace(vValue, "\b", vbBack)
                vValue = Replace(vValue, "\f", vbFormFeed)
                vValue = Replace(vValue, "\n", vbLf)
                vValue = Replace(vValue, "\r", vbCr)
                vValue = Replace(vValue, "\t", vbTab)
                vValue = Replace(vValue, "\[", "[")
                vValue = Replace(vValue, "\{", "{")
                vValue = Replace(vValue, "\}", "}")
                vValue = Replace(vValue, "\]", "]")
                'based on a performance comparison of the following Like pattern match
                '   against oRE.Test(vValue), the LIKE was just a little bit faster
                'If oRE.test(vValue) Then
                If vValue Like "*\[Uu][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]*" Then
                    Set oUMatches = oRE.Execute(vValue)
                    For Each oUM In oUMatches
                        vValue = Replace(vValue, oUM, ChrW("&h" & oUM.submatches(0)))
                    Next
                End If
            Else
                Select Case True
                    Case vValue Like "####-##-##T##:##:##Z", vValue Like "####-##-##T##:##:##.#*Z"
                        vValue = Replace(vValue, "T", " ")
                        vValue = Left(vValue, 19)
                        vValue = CDate(vValue)
                    Case vValue Like "####-##-##T##:##:##[+-]###"
                        lngZoffset = Right(vValue, 4)
                        vValue = Replace(vValue, "T", " ")
                        vValue = Left(vValue, 19)
                        vValue = DateAdd("n", lngZoffset, CDate(vValue))
                    Case vValue Like "####-##-##T##:##:##[+-]##:##"
                        dtZoffset = Right(vValue, 6)
                        vValue = Replace(vValue, "T", " ")
                        vValue = Left(vValue, 19)
                        vValue = dtZoffset + CDate(vValue)
                End Select
            End If
        Else
            vValue = oM.submatches(oM.submatches.Count - 1)      'lngObjectType)
            'cast the non-string parsed value into a VB value
            Select Case vValue
                Case vbNullString   'value will be an object
                Case "null"
                    vValue = Null
                Case "true"
                    vValue = True
                Case "false"
                    vValue = False
                Case Else
                    Select Case True
                        Case IsDate(vValue)
                            vValue = CDate(vValue)
                        Case IsNumeric(vValue)
                            vValue = Val(vValue)
                    End Select
            End Select
        End If
        
        
        Select Case lngObjectType
            Case 3      'Dictionary
                parmCurObj(oM.submatches(0)) = vValue
            Case 1      'Collection
                parmCurObj.Add vValue
        End Select
    Next
End Sub

Sub PushObjectOntoStack(parmObject As Object, parmStack As Collection)
    If parmStack.Count = 0 Then
        parmStack.Add parmObject
    
    Else    'push on top of stack (eg. parmObject becomes the 1st item)
        parmStack.Add parmObject, , 1
    End If
End Sub


