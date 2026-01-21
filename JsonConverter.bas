Attribute VB_Name = "JsonConverter"
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
' JSON Converter for VBA

Option Explicit

Private Const JSON_QUOTE As String = """"

Public Function ParseJson(ByVal jsonString As String) As Object
    Dim index As Long
    index = 1

    ' Skip whitespace
    SkipWhitespace jsonString, index

    Set ParseJson = ParseValue(jsonString, index)
End Function

Private Function ParseValue(ByRef jsonString As String, ByRef index As Long) As Variant
    Dim char As String

    SkipWhitespace jsonString, index

    If index > Len(jsonString) Then
        ParseValue = Null
        Exit Function
    End If

    char = Mid$(jsonString, index, 1)

    Select Case char
        Case "{"
            Set ParseValue = ParseObject(jsonString, index)
        Case "["
            Set ParseValue = ParseArray(jsonString, index)
        Case JSON_QUOTE
            ParseValue = ParseString(jsonString, index)
        Case "t", "f"
            ParseValue = ParseBoolean(jsonString, index)
        Case "n"
            ParseValue = ParseNull(jsonString, index)
        Case Else
            If char = "-" Or IsNumeric(char) Then
                ParseValue = ParseNumber(jsonString, index)
            Else
                Err.Raise 10001, "JsonConverter", "Invalid JSON at position " & index
            End If
    End Select
End Function

Private Function ParseObject(ByRef jsonString As String, ByRef index As Long) As Object
    Dim dict As Object
    Dim key As String
    Dim char As String

    Set dict = CreateObject("Scripting.Dictionary")

    ' Skip opening brace
    index = index + 1
    SkipWhitespace jsonString, index

    ' Check for empty object
    If Mid$(jsonString, index, 1) = "}" Then
        index = index + 1
        Set ParseObject = dict
        Exit Function
    End If

    Do
        SkipWhitespace jsonString, index

        ' Parse key
        If Mid$(jsonString, index, 1) <> JSON_QUOTE Then
            Err.Raise 10002, "JsonConverter", "Expected string key at position " & index
        End If
        key = ParseString(jsonString, index)

        SkipWhitespace jsonString, index

        ' Skip colon
        If Mid$(jsonString, index, 1) <> ":" Then
            Err.Raise 10003, "JsonConverter", "Expected colon at position " & index
        End If
        index = index + 1

        SkipWhitespace jsonString, index

        ' Parse value
        If IsObject(ParseValue(jsonString, index)) Then
            Set dict(key) = ParseValue(jsonString, index - 1)
            index = index - 1
            Dim tempVal As Variant
            Set tempVal = ParseValue(jsonString, index)
            Set dict(key) = tempVal
        Else
            dict(key) = ParseValue(jsonString, index)
        End If

        SkipWhitespace jsonString, index

        char = Mid$(jsonString, index, 1)
        If char = "}" Then
            index = index + 1
            Exit Do
        ElseIf char = "," Then
            index = index + 1
        Else
            Err.Raise 10004, "JsonConverter", "Expected comma or closing brace at position " & index
        End If
    Loop

    Set ParseObject = dict
End Function

Private Function ParseArray(ByRef jsonString As String, ByRef index As Long) As Object
    Dim arr As Object
    Dim char As String
    Dim val As Variant

    Set arr = CreateObject("System.Collections.ArrayList")

    ' Skip opening bracket
    index = index + 1
    SkipWhitespace jsonString, index

    ' Check for empty array
    If Mid$(jsonString, index, 1) = "]" Then
        index = index + 1
        Set ParseArray = ConvertToCollection(arr)
        Exit Function
    End If

    Do
        SkipWhitespace jsonString, index

        ' Parse value
        val = ParseValue(jsonString, index)
        If IsObject(val) Then
            arr.Add val
        Else
            arr.Add val
        End If

        SkipWhitespace jsonString, index

        char = Mid$(jsonString, index, 1)
        If char = "]" Then
            index = index + 1
            Exit Do
        ElseIf char = "," Then
            index = index + 1
        Else
            Err.Raise 10005, "JsonConverter", "Expected comma or closing bracket at position " & index
        End If
    Loop

    Set ParseArray = ConvertToCollection(arr)
End Function

Private Function ConvertToCollection(arr As Object) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim item As Variant

    For i = 0 To arr.Count - 1
        item = arr(i)
        If IsObject(item) Then
            col.Add item
        Else
            col.Add item
        End If
    Next i

    Set ConvertToCollection = col
End Function

Private Function ParseString(ByRef jsonString As String, ByRef index As Long) As String
    Dim result As String
    Dim char As String
    Dim escaped As Boolean

    result = ""

    ' Skip opening quote
    index = index + 1

    Do While index <= Len(jsonString)
        char = Mid$(jsonString, index, 1)

        If escaped Then
            Select Case char
                Case JSON_QUOTE, "\", "/"
                    result = result & char
                Case "b"
                    result = result & vbBack
                Case "f"
                    result = result & vbFormFeed
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    result = result & ChrW(CLng("&H" & Mid$(jsonString, index + 1, 4)))
                    index = index + 4
            End Select
            escaped = False
        ElseIf char = "\" Then
            escaped = True
        ElseIf char = JSON_QUOTE Then
            index = index + 1
            Exit Do
        Else
            result = result & char
        End If

        index = index + 1
    Loop

    ParseString = result
End Function

Private Function ParseNumber(ByRef jsonString As String, ByRef index As Long) As Variant
    Dim startIndex As Long
    Dim char As String
    Dim numStr As String

    startIndex = index

    ' Handle negative
    If Mid$(jsonString, index, 1) = "-" Then
        index = index + 1
    End If

    ' Integer part
    Do While index <= Len(jsonString)
        char = Mid$(jsonString, index, 1)
        If Not IsNumeric(char) Then Exit Do
        index = index + 1
    Loop

    ' Decimal part
    If index <= Len(jsonString) And Mid$(jsonString, index, 1) = "." Then
        index = index + 1
        Do While index <= Len(jsonString)
            char = Mid$(jsonString, index, 1)
            If Not IsNumeric(char) Then Exit Do
            index = index + 1
        Loop
    End If

    ' Exponent part
    If index <= Len(jsonString) Then
        char = Mid$(jsonString, index, 1)
        If char = "e" Or char = "E" Then
            index = index + 1
            If index <= Len(jsonString) Then
                char = Mid$(jsonString, index, 1)
                If char = "+" Or char = "-" Then
                    index = index + 1
                End If
            End If
            Do While index <= Len(jsonString)
                char = Mid$(jsonString, index, 1)
                If Not IsNumeric(char) Then Exit Do
                index = index + 1
            Loop
        End If
    End If

    numStr = Mid$(jsonString, startIndex, index - startIndex)

    If InStr(numStr, ".") > 0 Or InStr(numStr, "e") > 0 Or InStr(numStr, "E") > 0 Then
        ParseNumber = CDbl(numStr)
    Else
        On Error Resume Next
        ParseNumber = CLng(numStr)
        If Err.Number <> 0 Then
            Err.Clear
            ParseNumber = CDbl(numStr)
        End If
        On Error GoTo 0
    End If
End Function

Private Function ParseBoolean(ByRef jsonString As String, ByRef index As Long) As Boolean
    If Mid$(jsonString, index, 4) = "true" Then
        ParseBoolean = True
        index = index + 4
    ElseIf Mid$(jsonString, index, 5) = "false" Then
        ParseBoolean = False
        index = index + 5
    Else
        Err.Raise 10006, "JsonConverter", "Invalid boolean at position " & index
    End If
End Function

Private Function ParseNull(ByRef jsonString As String, ByRef index As Long) As Variant
    If Mid$(jsonString, index, 4) = "null" Then
        ParseNull = Null
        index = index + 4
    Else
        Err.Raise 10007, "JsonConverter", "Invalid null at position " & index
    End If
End Function

Private Sub SkipWhitespace(ByRef jsonString As String, ByRef index As Long)
    Dim char As String
    Do While index <= Len(jsonString)
        char = Mid$(jsonString, index, 1)
        If char <> " " And char <> vbCr And char <> vbLf And char <> vbTab Then
            Exit Do
        End If
        index = index + 1
    Loop
End Sub
