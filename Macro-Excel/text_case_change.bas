Option Explicit

' ===================== PUBLIC ENTRY =====================
Sub TextChangeCase()
    Dim choice As String
    Dim exceptions As String, acronyms As String
    Dim keepAfterListCasing As VbMsgBoxResult
    Dim treatColonAsBoundary As VbMsgBoxResult
    Dim addSpaceAfterColon As VbMsgBoxResult
    Dim normalizeMarkers As VbMsgBoxResult
    Dim cells As Range, cell As Range
    Dim txt As String
    Dim excArr() As String, acrArr() As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cells first.", vbExclamation
        Exit Sub
    End If
    Set cells = Selection

    choice = InputBox( _
        "Choose case option:" & vbCrLf & _
        "1 = Sentence case" & vbCrLf & _
        "2 = lowercase" & vbCrLf & _
        "3 = UPPERCASE" & vbCrLf & _
        "4 = Proper Case (Title Case)", _
        "Change Text Case")
    If choice = "" Then Exit Sub
    
    If choice = "1" Then
        exceptions = InputBox("Enter words to keep lowercase (comma-separated)" & vbCrLf & _
                              "Example: i.e.,vs.,etc.", "Sentence Case: Lowercase Exceptions")
        If exceptions = "" Then Exit Sub
        If Len(Trim$(exceptions)) > 0 Then
            excArr = SplitToCleanArray(LCase$(exceptions))
        Else
            Erase excArr
        End If
        
        acronyms = InputBox("Enter ACRONYMS to keep uppercase (comma-separated)" & vbCrLf & _
                            "Example: CEO,USA,UN", "Sentence Case: Acronyms to Preserve")
        If acronyms = "" Then Exit Sub
        If Len(Trim$(acronyms)) > 0 Then
            acrArr = SplitToCleanArray(UCase$(acronyms))
        Else
            Erase acrArr
        End If
        
        keepAfterListCasing = MsgBox( _
            "After list markers (e.g., 1., a), I.)," & vbCrLf & _
            "do you want to PRESERVE the original casing of the remainder?" & vbCrLf & vbCrLf & _
            "Yes  = Only uppercase the first letter after the marker; keep rest as-is." & vbCrLf & _
            "No   = Apply standard sentence case (lowercase the rest).", _
            vbYesNoCancel + vbQuestion, "List Marker Handling")
            If keepAfterListCasing = vbCancel Then Exit Sub
        
        treatColonAsBoundary = MsgBox( _
            "Treat the colon ':' as a sentence boundary?" & vbCrLf & _
            "Yes = Capitalize first word after ':'", _
            vbYesNoCancel + vbQuestion, "Colon Behavior")
            If treatColonAsBoundary = vbCancel Then Exit Sub
        
        addSpaceAfterColon = MsgBox( _
            "Normalize spacing after ':' ?" & vbCrLf & _
            "Yes = ensure exactly one space after ':'", _
            vbYesNoCancel + vbQuestion, "Colon Spacing")
            If addSpaceAfterColon = vbCancel Then Exit Sub
        
        normalizeMarkers = MsgBox( _
            "Normalize list-marker spacing at the start?" & vbCrLf & _
            "Examples normalized:" & vbCrLf & _
            "  1.a  ->  1. a" & vbCrLf & _
            "  1 )a ->  1) a" & vbCrLf & _
            "  -a   ->  - a" & vbCrLf & _
            "Yes = normalize before casing.", _
            vbYesNoCancel + vbQuestion, "List Marker Spacing")
            If normalizeMarkers = vbCancel Then Exit Sub
    End If
    
    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    
    For Each cell In cells
        If VarType(cell.Value) = vbString Then
            txt = CStr(cell.Value)
            
            ' Optional marker normalization
            If choice = "1" And normalizeMarkers = vbYes Then
                txt = NormalizeListMarkerSpacing(txt)
            End If
            
            Select Case choice
                Case "1"
                    txt = SentenceCase( _
                              txt, _
                              excArr, _
                              acrArr, _
                              (keepAfterListCasing = vbYes), _
                              (treatColonAsBoundary = vbYes), _
                              (addSpaceAfterColon = vbYes))
                Case "2"
                    txt = LCase$(txt)
                Case "3"
                    txt = UCase$(txt)
                Case "4"
                    txt = Application.WorksheetFunction.Proper(txt)
                Case Else
                    MsgBox "Invalid choice. Enter 1, 2, 3, or 4.", vbExclamation
                    GoTo CleanFail
            End Select
            cell.Value = txt
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "Case change applied!", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Stopped due to an error: " & Err.Description, vbExclamation
End Sub

' ===================== CORE SENTENCE CASE =====================
Private Function SentenceCase(ByVal s As String, _
                                   ByRef exceptions() As String, _
                                   ByRef acronyms() As String, _
                                   ByVal preserveAfterList As Boolean, _
                                   ByVal colonAsBoundary As Boolean, _
                                   ByVal spaceAfterColon As Boolean) As String
    Dim parts As Collection, out As String
    Dim i As Long, segment As String, delim As String
    
    If spaceAfterColon Then s = NormalizeColonSpacing(s)
    Set parts = SplitIntoSegmentsKeepDelims(s, colonAsBoundary)
    
    For i = 1 To parts.Count Step 2
        segment = parts(i)
        If i + 1 <= parts.Count Then
            delim = parts(i + 1)
        Else
            delim = ""
        End If
        
        segment = ApplyListAwareSentenceCase(segment, exceptions, acronyms, preserveAfterList)
        out = out & segment & delim
    Next i
    
    SentenceCase = out
End Function

Private Function SplitIntoSegmentsKeepDelims(ByVal s As String, ByVal colonBoundary As Boolean) As Collection
    Dim re As Object, matches As Object, m As Object
    Dim result As New Collection
    Dim startPos As Long, pattern As String
    
    pattern = "([\.!\?;]|[\r\n]+)"
    If colonBoundary Then pattern = "([\.!\?:;]|[\r\n]+)"
    
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = pattern
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
    End With
    
    startPos = 1
    Set matches = re.Execute(s)
    If matches.Count = 0 Then
        result.Add s
        Exit Function
    End If
    
    For Each m In matches
        result.Add Mid$(s, startPos, m.FirstIndex + 1 - startPos) ' text before delimiter
        result.Add m.Value                                        ' delimiter
        startPos = m.FirstIndex + m.Length + 1
    Next m
    If startPos <= Len(s) Then result.Add Mid$(s, startPos)
    
    Set SplitIntoSegmentsKeepDelims = result
End Function

Private Function ApplyListAwareSentenceCase(ByVal seg As String, _
                                            ByRef exceptions() As String, _
                                            ByRef acronyms() As String, _
                                            ByVal preserveAfterList As Boolean) As String
    Dim re As Object, m As Object
    Dim rest As String, prefix As String
    
    ' Detect list marker at start (number/letter/roman) then '.' or ')', allowing spaces
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "^\s*((?:\d+|[A-Za-z]|[ivxlcdmIVXLCDM]+)\s*[\.\)]\s*)"  ' e.g., "1 . )   "
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
    End With
    
    If re.Test(seg) Then
        Set m = re.Execute(seg)(0)
        prefix = Left$(seg, m.FirstIndex + m.Length)
        rest = Mid$(seg, m.FirstIndex + m.Length + 1)
        rest = CapitalizeFirstAlphaAfterSpaces(rest, preserveAfterList)
        rest = SentenceCaseOneSegment(rest, exceptions, acronyms, preserveAfterList)
        ApplyListAwareSentenceCase = prefix & rest
    Else
        ApplyListAwareSentenceCase = SentenceCaseOneSegment( _
            CapitalizeFirstAlphaAfterSpaces(seg, False), _
            exceptions, acronyms, False)
    End If
End Function

' Apply sentence-case to a single segment (no delimiter inside)
Private Function SentenceCaseOneSegment(ByVal s As String, _
                                        ByRef exceptions() As String, _
                                        ByRef acronyms() As String, _
                                        ByVal keepRest As Boolean) As String
    Dim tokens As Collection, t As Variant, out As String
    Dim firstDone As Boolean, part As String

    Set tokens = TokenizeWordsAndSeparators(s)

    For Each t In tokens
        part = CStr(t)

        If Not firstDone Then
            If IsAlpha(part) Then
                ' If the first word is an acronym, keep it fully UPPERCASE
                If ArrayContains(acronyms, UCase$(part)) Then
                    out = out & UCase$(part)
                Else
                    ' Otherwise sentence-case the first word
                    If keepRest Then
                        out = out & UCase$(Left$(part, 1)) & Mid$(part, 2)
                    Else
                        out = out & UCase$(Left$(part, 1)) & LCase$(Mid$(part, 2))
                    End If
                End If
                firstDone = True
            Else
                out = out & part
            End If
        Else
            ' After first word
            out = out & LowercaseWordRespectAcronyms(part, acronyms, Not keepRest)
        End If
    Next t

    ' Lowercase exceptions as whole words
    out = ApplyLowercaseExceptions(out, exceptions)
    SentenceCaseOneSegment = out
End Function


' ===================== NORMALIZERS =====================
Private Function NormalizeListMarkerSpacing(ByVal s As String) As String
    Dim re As Object
    
    ' 1.a / 1 .a / 1 . a / 1) a / 1 )a / 1 ) a  ->  "1. a" or "1) a"
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "^\s*(\d+|[A-Za-z]|[ivxlcdmIVXLCDM]+)\s*([\.\)])\s*(.*)$"
        .Global = False
        .IgnoreCase = True
    End With
    If re.Test(s) Then
        Dim m As Object: Set m = re.Execute(s)(0)
        NormalizeListMarkerSpacing = Trim$(m.SubMatches(0)) & m.SubMatches(1) & " " & Trim$(m.SubMatches(2))
        Exit Function
    End If
    
    ' Bullets: -a / -  a / •a -> "- a" / "• a"
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "^\s*([\-•])\s*(.+)$"
        .Global = False
        .IgnoreCase = True
    End With
    If re.Test(s) Then
        Dim m2 As Object: Set m2 = re.Execute(s)(0)
        NormalizeListMarkerSpacing = m2.SubMatches(0) & " " & Trim$(m2.SubMatches(1))
        Exit Function
    End If
    
    NormalizeListMarkerSpacing = s
End Function

Private Function NormalizeColonSpacing(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = ":\s*"
        .Global = True
    End With
    NormalizeColonSpacing = re.Replace(s, ": ")
End Function

' ===================== TOKEN HELPERS =====================
Private Function TokenizeWordsAndSeparators(ByVal s As String) As Collection
    Dim re As Object, m As Object, result As New Collection
    Dim lastPos As Long: lastPos = 1
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "([A-Za-z]+)"
        .Global = True
    End With
    For Each m In re.Execute(s)
        If m.FirstIndex + 1 > lastPos Then
            result.Add Mid$(s, lastPos, m.FirstIndex + 1 - lastPos)
        End If
        result.Add m.Value
        lastPos = m.FirstIndex + m.Length + 1
    Next m
    If lastPos <= Len(s) Then result.Add Mid$(s, lastPos)
    Set TokenizeWordsAndSeparators = result
End Function

Private Function CapitalizeFirstAlpha(ByVal s As String, ByVal keepRest As Boolean) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            If keepRest Then
                CapitalizeFirstAlpha = Left$(s, i - 1) & UCase$(ch) & Mid$(s, i + 1)
            Else
                CapitalizeFirstAlpha = Left$(s, i - 1) & UCase$(ch) & LCase$(Mid$(s, i + 1))
            End If
            Exit Function
        End If
    Next i
    CapitalizeFirstAlpha = s
End Function

Private Function LowercaseWordRespectAcronyms(ByVal part As String, _
                                              ByRef acronyms() As String, _
                                              ByVal lowerRest As Boolean) As String
    If IsAlpha(part) Then
        If ArrayContains(acronyms, UCase$(part)) Then
            LowercaseWordRespectAcronyms = UCase$(part)
        Else
            If lowerRest Then
                LowercaseWordRespectAcronyms = LCase$(part)
            Else
                LowercaseWordRespectAcronyms = part
            End If
        End If
    Else
        LowercaseWordRespectAcronyms = part
    End If
End Function

Private Function IsAlpha(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "^[A-Za-z]+$"
        .Global = False
    End With
    IsAlpha = re.Test(s)
End Function

Private Function CapitalizeFirstAlphaAfterSpaces(ByVal s As String, ByVal keepRest As Boolean) As String
    Dim i As Long, ch As String
    ' skip any leading spaces/tabs/etc.
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            If keepRest Then
                CapitalizeFirstAlphaAfterSpaces = Left$(s, i - 1) & UCase$(ch) & Mid$(s, i + 1)
            Else
                CapitalizeFirstAlphaAfterSpaces = Left$(s, i - 1) & UCase$(ch) & LCase$(Mid$(s, i + 1))
            End If
            Exit Function
        End If
        i = i + 1
    Loop
    CapitalizeFirstAlphaAfterSpaces = s
End Function

' ===================== REPLACERS & UTILS =====================
Private Function ApplyLowercaseExceptions(ByVal s As String, ByRef exceptions() As String) As String
    Dim i As Long, ex As String
    If Not ArrayHasItems(exceptions) Then
        ApplyLowercaseExceptions = s
        Exit Function
    End If
    For i = LBound(exceptions) To UBound(exceptions)
        ex = Trim$(exceptions(i))
        If Len(ex) > 0 Then s = ReplaceTokenCaseInsensitive(s, ex, LCase$(ex))
    Next i
    ApplyLowercaseExceptions = s
End Function

Private Function ReplaceTokenCaseInsensitive(ByVal s As String, ByVal findTok As String, ByVal repl As String) As String
    Dim re As Object, pat As String
    Set re = CreateObject("VBScript.RegExp")

    ' Escape regex metacharacters in the token
    findTok = ReplaceRegexEscape(findTok)

    ' Left boundary captured (start or non-alnum), token, right boundary via look-ahead
    ' (^|[^A-Za-z0-9]) (token) (?=$|[^A-Za-z0-9])
    pat = "(^|[^A-Za-z0-9])(" & findTok & ")(?=$|[^A-Za-z0-9])"

    With re
        .Pattern = pat
        .Global = True
        .IgnoreCase = True
    End With

    ' Keep the left boundary, replace just the token
    ReplaceTokenCaseInsensitive = re.Replace(s, "$1" & repl)
End Function

Private Function ReplaceRegexEscape(ByVal s As String) As String
    Dim ch As String, i As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("\^$.|?*+()[]{}", ch) > 0 Then
            out = out & "\" & ch
        Else
            out = out & ch
        End If
    Next i
    ReplaceRegexEscape = out
End Function

Private Function SplitToCleanArray(ByVal csv As String) As String()
    Dim parts() As String, i As Long
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(parts(i))
    Next i
    SplitToCleanArray = parts
End Function

Private Function ArrayHasItems(ByRef arr() As String) As Boolean
    On Error Resume Next
    ArrayHasItems = (LBound(arr) <= UBound(arr))
    On Error GoTo 0
End Function

Private Function ArrayContains(ByRef arr() As String, ByVal item As String) As Boolean
    Dim i As Long
    If Not ArrayHasItems(arr) Then Exit Function
    For i = LBound(arr) To UBound(arr)
        If StrComp(Trim$(arr(i)), item, vbTextCompare) = 0 Then
            ArrayContains = True
            Exit Function
        End If
    Next i
End Function