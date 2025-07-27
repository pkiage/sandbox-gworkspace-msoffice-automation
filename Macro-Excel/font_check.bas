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

Option Explicit

'==============================================================
'  PUBLIC ENTRY POINT
'==============================================================
Sub AuditFonts()
    Dim dictDetail As Object, dictSummary As Object
    Set dictDetail = CreateObject("Scripting.Dictionary")
    Set dictSummary = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet, rng As Range, c As Range
    
    '─── Walk every sheet ──────────────────────────────────────
    For Each ws In ActiveWorkbook.Worksheets
        
        'Cells in UsedRange
        On Error Resume Next
        Set rng = ws.UsedRange
        On Error GoTo 0
        
        If Not rng Is Nothing Then
            For Each c In rng.Cells
                RecordFont dictDetail, dictSummary, _
                           c.Font.Name, c.Font.Size, c.Font.Color, _
                           "Cell", ws.Name & "!" & c.Address(False, False)
            Next c
        End If
        
        'Embedded charts
        Dim chObj As ChartObject
        For Each chObj In ws.ChartObjects
            HandleChartFonts chObj.Chart, dictDetail, dictSummary, _
                              ws.Name & "!" & chObj.Name
        Next chObj
    Next ws
    
    '─── Dump the dictionaries to sheets ───────────────────────
    DumpDetail dictDetail
    DumpSummary dictSummary
    MsgBox "Font audit complete – see 'Font Audit' and 'Font Summary'.", vbInformation
End Sub

'==============================================================
'  HELPERS
'==============================================================
'Return safe text for Name / Size
Private Function SafeText(v As Variant, Optional emptyToken As String = "(mixed)") As String
    If IsError(v) Or IsMissing(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeText = emptyToken
    Else
        SafeText = Trim$(CStr(v))
        If Len(SafeText) = 0 Then SafeText = emptyToken
    End If
End Function

'Return safe RGB hex for Color
Private Function SafeColor(v As Variant, Optional emptyToken As String = "(mixed)") As String
    If IsError(v) Or IsMissing(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeColor = emptyToken
    Else
        'Keep only the low 24 bits then format as #RRGGBB
        SafeColor = "#" & Right$("000000" & Hex$(v Mod &H1000000), 6)
    End If
End Function

'Add a record to both dictionaries
Private Sub RecordFont(dictDet As Object, dictSum As Object, _
                       fName As Variant, fSize As Variant, fColor As Variant, _
                       objType As String, loc As String)
    
    Dim n As String: n = SafeText(fName)
    Dim s As String: s = SafeText(fSize)
    Dim col As String: col = SafeColor(fColor)
    
    'Detail key → Name|Size|Color|Object|Location
    Dim kDet As String: kDet = n & "|" & s & "|" & col & "|" & objType & "|" & loc
    dictDet(kDet) = dictDet(kDet) + 1     'Using late‑bound default .Item
    
    'Summary key → Name|Size|Color
    Dim kSum As String: kSum = n & "|" & s & "|" & col
    dictSum(kSum) = dictSum(kSum) + 1
End Sub

'Scan chart parts
Private Sub HandleChartFonts(cht As Chart, _
                             dictDet As Object, dictSum As Object, baseLoc As String)
    Dim ax As Axis, s As Series
    
    'Chart title
    If cht.HasTitle Then _
        RecordFont dictDet, dictSum, _
                   cht.ChartTitle.Font.Name, cht.ChartTitle.Font.Size, cht.ChartTitle.Font.Color, _
                   "ChartTitle", baseLoc
    
    'Legend
    If cht.HasLegend Then _
        RecordFont dictDet, dictSum, _
                   cht.Legend.Font.Name, cht.Legend.Font.Size, cht.Legend.Font.Color, _
                   "Legend", baseLoc
    
    'Axes – titles + tick labels
    For Each ax In cht.Axes
        On Error Resume Next
        If ax.HasTitle Then _
            RecordFont dictDet, dictSum, _
                       ax.AxisTitle.Font.Name, ax.AxisTitle.Font.Size, ax.AxisTitle.Font.Color, _
                       "AxisTitle", baseLoc & " (" & AxisName(ax.Type) & ")"
        
        RecordFont dictDet, dictSum, _
                   ax.TickLabels.Font.Name, ax.TickLabels.Font.Size, ax.TickLabels.Font.Color, _
                   "AxisLabels", baseLoc & " (" & AxisName(ax.Type) & ")"
        On Error GoTo 0
    Next ax
    
    'Series data‑labels (first label as proxy)
    For Each s In cht.SeriesCollection
        If s.HasDataLabels Then
            On Error Resume Next
            RecordFont dictDet, dictSum, _
                       s.DataLabels(1).Font.Name, s.DataLabels(1).Font.Size, s.DataLabels(1).Font.Color, _
                       "DataLabel", baseLoc & " (" & s.Name & ")"
            On Error GoTo 0
        End If
    Next s
End Sub

'Axis type → friendly text
Private Function AxisName(axType As XlAxisType) As String
    Select Case axType
        Case xlCategory:     AxisName = "Category"
        Case xlValue:        AxisName = "Value"
        Case xlSeriesAxis:   AxisName = "Series"
        Case Else:           AxisName = "Axis"
    End Select
End Function

'Write detail sheet
Private Sub DumpDetail(d As Object)
    Const shtName As String = "Font Audit"
    ReplaceSheet shtName
    Dim outSht As Worksheet: Set outSht = Worksheets(shtName)
    
    outSht.Range("A1:F1").Value = _
        Array("Font Name", "Size", "Color", "Object Type", "Location", "Occurrences")
    
    Dim r As Long: r = 2
    Dim k As Variant, parts
    For Each k In d.Keys
        parts = Split(k, "|")                      'Name,Size,Color,Object,Location
        outSht.Cells(r, 1).Resize(1, 5).Value = parts
        outSht.Cells(r, 6).Value = d(k)
        r = r + 1
    Next k
    outSht.Columns.AutoFit
End Sub

'Write summary sheet
Private Sub DumpSummary(d As Object)
    Const shtName As String = "Font Summary"
    ReplaceSheet shtName
    Dim outSht As Worksheet: Set outSht = Worksheets(shtName)
    
    outSht.Range("A1:D1").Value = _
        Array("Font Name", "Size", "Color", "Occurrences")
    
    Dim r As Long: r = 2
    Dim k As Variant, parts
    For Each k In d.Keys
        parts = Split(k, "|")                      'Name,Size,Color
        outSht.Cells(r, 1).Resize(1, 3).Value = parts
        outSht.Cells(r, 4).Value = d(k)
        r = r + 1
    Next k
    outSht.Columns.AutoFit
End Sub

'Delete sheet if it exists, then create a fresh one
Private Sub ReplaceSheet(shtName As String)
    Application.DisplayAlerts = False
    On Error Resume Next: Worksheets(shtName).Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    Worksheets.Add(Before:=Sheets(1)).Name = shtName
End Sub
