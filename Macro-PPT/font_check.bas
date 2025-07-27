' Copyright 2025 The Contributors
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

'==================  Font Family × Size Lister  ==================
' Same output: reports every font and the sizes it appears in
'=================================================================
Option Explicit

'──────── MAIN ────────
Sub ListFontsAndSizesInDeck()

    Dim pres As Presentation: Set pres = ActivePresentation
    Dim fonts As Object: Set fonts = CreateObject("Scripting.Dictionary")
    
    Dim sld As Slide, shp As Shape
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            CollectFontSizeInfo shp, sld.SlideIndex, fonts
        Next shp
    Next sld
    
    If fonts.Count = 0 Then
        MsgBox "No text found in this presentation.", vbInformation
        Exit Sub
    End If
    
    Dim fontNames() As String: fontNames = DictKeysSortedAlpha(fonts)
    CreateFontSizeReportSlide pres, fontNames, fonts
    
    MsgBox "Font & size report created at the end of the deck.", _
           vbInformation, "Audit complete"
End Sub


'──────── RECURSIVE COLLECTOR ──────── 
Private Sub CollectFontSizeInfo(ByVal shp As Shape, ByVal slideNo As Long, _
                                ByRef fonts As Object)
                                
    Dim i&, j&, k&, tr As TextRange
    
    Select Case True
    'text boxes & placeholders
    Case shp.HasTextFrame
        If shp.TextFrame.HasText Then
            Set tr = shp.TextFrame.TextRange
            For i = 1 To tr.Runs.Count
                AddFontRecord fonts, _
                              tr.Runs(i).Font.Name, _
                              Round(tr.Runs(i).Font.Size, 2), _
                              slideNo
            Next i
        End If
        
    'tables
    Case shp.HasTable
        For i = 1 To shp.Table.Rows.Count
            For j = 1 To shp.Table.Columns.Count
                With shp.Table.Cell(i, j).Shape.TextFrame.TextRange
                    For k = 1 To .Runs.Count
                        AddFontRecord fonts, _
                                      .Runs(k).Font.Name, _
                                      Round(.Runs(k).Font.Size, 2), _
                                      slideNo
                    Next k
                End With
            Next j
        Next i
        
    'groups – recurse
    Case shp.Type = msoGroup
        For i = 1 To shp.GroupItems.Count
            CollectFontSizeInfo shp.GroupItems(i), slideNo, fonts
        Next i
    End Select
End Sub


'──────── ADD / UPDATE FONT‑SIZE RECORD ──────── 
Private Sub AddFontRecord(ByRef fonts As Object, _
                           ByVal fName As String, _
                           ByVal fSize As Double, _
                           ByVal slideNo As Long)
                           
    Dim sizesDict As Object, infoArr, slidesDict As Object
    Dim sizeKey$: sizeKey = CStr(fSize)
    
    If Not fonts.exists(fName) Then _
        fonts.Add fName, CreateObject("Scripting.Dictionary")
    Set sizesDict = fonts(fName)
    
    If Not sizesDict.exists(sizeKey) Then
        Set slidesDict = CreateObject("Scripting.Dictionary")
        infoArr = Array(0&, slidesDict)           'count, slidesDict
        sizesDict.Add sizeKey, infoArr
    Else
        infoArr = sizesDict(sizeKey)
        Set slidesDict = infoArr(1)
    End If
    
    infoArr(0) = CLng(infoArr(0)) + 1
    If Not slidesDict.exists(slideNo) Then slidesDict.Add slideNo, True
    sizesDict(sizeKey) = infoArr                 'write back
End Sub


'──────── REPORT SLIDE CREATOR ────────
Private Sub CreateFontSizeReportSlide(ByRef pres As Presentation, _
                                      ByRef fontNames() As String, _
                                      ByRef fonts As Object)
    Const TITLE$ = "Font & Size Usage Report"
    Dim rpt As Slide: Set rpt = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutText)
    rpt.MoveTo pres.Slides.Count
    rpt.Shapes.Title.TextFrame.TextRange.Text = TITLE
    
    Dim bodyBox As Shape, ph As Shape
    For Each ph In rpt.Shapes.Placeholders
        If ph.PlaceholderFormat.Type = ppPlaceholderBody Then Set bodyBox = ph
    Next ph
    If bodyBox Is Nothing Then _
        Set bodyBox = rpt.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 100, 600, 350)
    
    Dim txt$, i&, fName$, sizesDict As Object, sizeKeys() As Double, j&, sKey$
    Dim infoArr, slidesDict As Object
    For i = LBound(fontNames) To UBound(fontNames)
        fName = fontNames(i)
        txt = txt & fName & ":" & vbCrLf
        Set sizesDict = fonts(fName)
        sizeKeys = DictKeysSortedNumeric(sizesDict)
        For j = LBound(sizeKeys) To UBound(sizeKeys)
            sKey = CStr(sizeKeys(j))
            infoArr = sizesDict(sKey): Set slidesDict = infoArr(1)
            txt = txt & "   • " & Format(sizeKeys(j), "0.##") & " pt" & _
                  " – " & infoArr(0) & " run(s), slide(s): " & _
                  JoinDictKeys(slidesDict) & vbCrLf
        Next j
        txt = txt & vbCrLf
    Next i
    
    bodyBox.TextFrame.TextRange.Text = txt
    On Error Resume Next: bodyBox.TextFrame.AutoSize = ppAutoSizeShapeToFitText
End Sub


'──────── SORT / JOIN HELPERS ────────
Private Function DictKeysSortedAlpha(dic As Object) As String()
    Dim arr() As String, i&, j&, tmp$
    ReDim arr(0 To dic.Count - 1)
    For i = 0 To dic.Count - 1: arr(i) = CStr(dic.Keys()(i)): Next i
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then _
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    DictKeysSortedAlpha = arr
End Function

Private Function DictKeysSortedNumeric(dic As Object) As Double()
    Dim arr() As Double, i&, j&, tmp#
    ReDim arr(0 To dic.Count - 1)
    For i = 0 To dic.Count - 1: arr(i) = CDbl(dic.Keys()(i)): Next i
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    DictKeysSortedNumeric = arr
End Function

Private Function JoinDictKeys(d As Object) As String
    Dim k, arr() As String, i&
    ReDim arr(0 To d.Count - 1)
    For Each k In d.Keys: arr(i) = k: i = i + 1: Next k
    JoinDictKeys = Join(arr, ", ")
End Function
