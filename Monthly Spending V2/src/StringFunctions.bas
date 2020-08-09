Attribute VB_Name = "StringFunctions"
Option Explicit


' Name: CountCharsInStr
' Description: Counts the number of given characters that occur in a string.
' Params:
'       text - this is the string to search.
'       char - these are the characters to count.
'       compareMethod - vbBinaryCompare = case sensitive. vbTextCompare = not case sensitive.
' Return value: Returns the number of occurences of the given character.
' Example:  CountCharsInStr("Mary Had a little lamb","a",vbTextCompare) - counts the
'            number of "a" characters in the string.
' https://excelmacromastery.com/
' YouTube Video: Excel VBA - The Missing Strings Functions(https://youtu.be/ibnWo1suz0w)
Public Function CountCharsInStr(thisText As String _
                            , thisChar As String _
                            , Optional compareMethod As VbCompareMethod = vbBinaryCompare) As Long
     
    CountCharsInStr = (Len(thisText) _
                        - Len(replace(thisText, thisChar, "", compare:=compareMethod))) _
                            / Len(thisChar)
     
End Function


' Name: sprintf
' Description: Provides an easy way to format a string.
' Params:
'       text - this is the string to format. It contains a token to represent the position that will be filled.
'              Tokens are %1,%2,%3 etc.
'       varStrings - these are a list of the items to replace the tokens
' Example:  sprintf("Workbook is %1","Data.xlsm")
' Return value: Returns the newly formatted string.
' https://excelmacromastery.com/
' YouTube Video: Excel VBA - The Missing Strings Functions(https://youtu.be/ibnWo1suz0w)
Public Function sprintf(ByVal thisText As String, ParamArray varStrings() As Variant) As String
    
    Dim i As Long
    On Error GoTo eh
    For i = LBound(varStrings) To UBound(varStrings)
        ' Ensure current parameter is valid
        If TypeName(varStrings(i)) <> "String" Then
            Err.Raise 5, "Invalid item passed as token. The string is parameter no [" & CStr(i + 1) & "]." _
                & " Utils.Printf"
            GoTo Continue
        End If
        thisText = replace(thisText, "%" & CStr(i + 1), varStrings(i))
Continue:
    Next
    
    sprintf = thisText

Done:
    Exit Function
eh:
    MsgBox Err.Description & "Utils.Printf. "
End Function

' Note: Add Reference Microsoft VBScript Regular Expressions 5.5
' Name: CleanString
' Description:  Remove duplicate characters from a string
' Example:  "Mary    had a   little      lamb"
' becomes "Mary had a little lamb"
' Params:
'       DirtyString - this is the string to clean.
'       replaceValue - these are the characters to remove.
'       replaceWith - these are the values to replace them with.
' Return value: Returns the updated string.
' https://excelmacromastery.com/
' YouTube Video: Excel VBA - The Missing Strings Functions(https://youtu.be/ibnWo1suz0w)
Public Function CleanStringRegEx(ByVal DirtyString As String _
                            , ByVal replaceValue As String) As String
    
    If DirtyString = "" Or replaceValue = "" Then
        Err.Raise vbObjectError + 1, "CleanString" _
            , "One of the parameter strings was empty."
        
    End If
    
    Dim regEx As New RegExp
    regEx.Global = True
    ' Set the pattern to find multiple occurences
    regEx.Pattern = replaceValue & "+"
        
    ' replace the the multiple occurences with one occurence
    CleanStringRegEx = regEx.replace(DirtyString, replaceValue)
    
End Function


' Name: CleanStringUsingLoop
' Description:  Remove duplicate characters from a string
' Example:  "Mary    had a   little      lamb"
' becomes "Mary had a little lamb"
' Params:
'       DirtyString - this is the string to clean.
'       replaceValue - these are the characters to remove.
'       replaceWith - these are the values to replace them with.
' Return value: Returns the updated string.
' https://excelmacromastery.com/
' YouTube Video: Excel VBA - The Missing Strings Functions(https://youtu.be/ibnWo1suz0w)
Function CleanStringUsingLoop(ByVal Text As String, val As String) As String

    Do
        Text = replace(Text, String(2, val), val)
    Loop While InStr(Text, String(2, val)) > 0
    
    CleanStringUsingLoop = Text
    
End Function
