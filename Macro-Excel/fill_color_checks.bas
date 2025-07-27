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

'---------------------------------------------------------------
'  PUBLIC WRAPPERS
'---------------------------------------------------------------
Sub AuditFillsFull()
    'Full detail  +  grouped view
    AuditFillsFast False
End Sub

Sub AuditFillsSummary()
    'Grouped view only (fastest)
    AuditFillsFast True
End Sub


'---------------------------------------------------------------
'  CORE ROUTINE
'---------------------------------------------------------------
Sub AuditFillsFast(Optional ByVal SummaryOnly As Boolean = False)
    
    Dim t0 As Single: t0 = Timer
    
    Dim det As Object, sum As Object
    Set det = CreateObject("Scripting.Dictionary")
    Set sum = CreateObject("Scripting.Dictionary")
    
    'Speed settings
    Dim wasCalc As XlCalculation
    With Application
        wasCalc = .Calculation
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    'Scan every worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ScanSheetFills ws, det, sum, SummaryOnly
    Next ws
    
    'Output sheets
    DumpSummary sum
    If Not SummaryOnly Then DumpDetail det
    
    'Restore settings
    With Application
        .Calculation = wasCalc
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    MsgBox IIf(SummaryOnly, "Summary", "Full detail") & _
           " audit finished in " & Format$(Timer - t0, "0.00") & " s", vbInformation
End Sub


'---------------------------------------------------------------
'  WORKSHEET-LEVEL SCAN
'---------------------------------------------------------------
Private Sub ScanSheetFills(ws As Worksheet, det As Object, sum As Object, _
                           SummaryOnly As Boolean)
    
    Dim rngConst As Range, rngForm As Range, area As Range
    
    On Error Resume Next
    Set rngConst = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    Set rngForm = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    '-- Cells – constants
    If Not rngConst Is Nothing Then
        For Each area In rngConst.Areas
            ProcessCellArea area, ws.name, det, sum, SummaryOnly
        Next area
    End If
    '-- Cells – formulas
    If Not rngForm Is Nothing Then
        For Each area In rngForm.Areas
            ProcessCellArea area, ws.name, det, sum, SummaryOnly
        Next area
    End If
    
    '-- Worksheet shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Fill.Visible Then
            RecordColor det, sum, _
                         FillHex(shp.Fill.ForeColor.RGB), "ShapeFill", _
                         ws.name & "!" & shp.name, SummaryOnly
        End If
    Next shp
    
    '-- Embedded charts
    Dim chObj As ChartObject
    For Each chObj In ws.ChartObjects
        HandleChartColors chObj.Chart, det, sum, _
                           ws.name & "!" & chObj.name, SummaryOnly
    Next chObj
End Sub


'---------------------------------------------------------------
'  PROCESS A CELL AREA (skips cells with default "no fill")
'---------------------------------------------------------------
Private Sub ProcessCellArea(area As Range, shtName As String, _
                            det As Object, sum As Object, SummaryOnly As Boolean)
    
    Dim c As Range, hexCol As String
    For Each c In area.Cells
        If c.Interior.ColorIndex <> xlColorIndexNone Then   'skip blank fills
            hexCol = FillHex(c.Interior.Color)
            RecordColor det, sum, hexCol, "CellFill", _
                         shtName & "!" & c.Address(False, False), SummaryOnly
        End If
    Next c
End Sub
'---------------------------------------------------------------
'  CHART  FILLS
'---------------------------------------------------------------
Private Sub HandleChartColors(cht As Chart, det As Object, sum As Object, _
                              baseLoc As String, SummaryOnly As Boolean)

    'Chart & plot backgrounds
    RecordColor det, sum, FillHex(cht.ChartArea.Format.Fill.ForeColor.RGB), _
                "ChartAreaFill", baseLoc, SummaryOnly
    RecordColor det, sum, FillHex(cht.PlotArea.Format.Fill.ForeColor.RGB), _
                "PlotAreaFill", baseLoc, SummaryOnly

    'Title & legend fills
    If cht.HasTitle Then _
        RecordColor det, sum, FillHex(cht.ChartTitle.Format.Fill.ForeColor.RGB), _
                    "ChartTitleFill", baseLoc, SummaryOnly
    If cht.HasLegend Then _
        RecordColor det, sum, FillHex(cht.Legend.Format.Fill.ForeColor.RGB), _
                    "LegendFill", baseLoc, SummaryOnly

    'Series / point colours (every point, but simpler location string)
    Dim s As Series, p As Point, hexCol As String, seen As Object
    For Each s In cht.SeriesCollection
        Set seen = CreateObject("Scripting.Dictionary")
        
        For Each p In s.Points
            hexCol = GetPointColourHex(p)   'Fill?Line?"" helper
            
            If Len(hexCol) > 0 And Not seen.exists(hexCol) Then
                seen.Add hexCol, True
                
                RecordColor det, sum, hexCol, "SeriesFill", _
                    baseLoc & " (" & SafeSeriesName(s) & ")", SummaryOnly
            End If
        Next p
    Next s
End Sub

'Helper: try Fill first, then Line colour
Private Function GetPointColourHex(pt As Point) As String
    Dim clr As Variant
    On Error Resume Next          'some point types lack Fill
    clr = pt.Format.Fill.ForeColor.RGB
    If Err.Number <> 0 Then
        Err.Clear
        clr = pt.Format.Line.ForeColor.RGB
    End If
    On Error GoTo 0
    
    If IsEmpty(clr) Or IsError(clr) Then
        GetPointColourHex = ""
    Else
        GetPointColourHex = FillHex(clr)
    End If
End Function

'Return series .Name or fallback "Series #<index>"
Private Function SafeSeriesName(s As Series) As String
    On Error Resume Next
    Dim nm As String: nm = CStr(s.name)
    If Err.Number <> 0 Or Len(nm) = 0 Then nm = "Series #" & s.Index
    SafeSeriesName = nm
    On Error GoTo 0
End Function

'---------------------------------------------------------------
'  RECORD COLOUR INTO DICTIONARIES
'---------------------------------------------------------------
Private Sub RecordColor(det As Object, sum As Object, colHex As String, _
                        objType As String, loc As String, SummaryOnly As Boolean)
    
    sum(colHex) = sum(colHex) + 1
    
    If Not SummaryOnly Then
        det(colHex & "|" & objType & "|" & loc) = _
            det(colHex & "|" & objType & "|" & loc) + 1
    End If
End Sub


'---------------------------------------------------------------
'  COLOR LONG to #RRGGBB
'---------------------------------------------------------------
Private Function FillHex(v As Variant) As String
    If v = -4142 Or v = 0 Or IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        FillHex = "(none)"
    Else
        Dim r As Long, g As Long, b As Long
        r = v And &HFF                      'red   = lowest byte
        g = (v \ 256) And &HFF              'green = mid byte
        b = (v \ 65536) And &HFF            'blue  = high byte
        
        FillHex = "#" & Right$("0" & Hex$(r), 2) & _
                        Right$("0" & Hex$(g), 2) & _
                        Right$("0" & Hex$(b), 2)
    End If
End Function


'---------------------------------------------------------------
'  #RRGGBB  to  COLOR LONG  (for swatches)
'---------------------------------------------------------------
Private Function HexToLong(ByVal hexStr As Variant) As Long
    If Left$(hexStr, 1) = "#" Then hexStr = Mid$(hexStr, 2)
    If Len(hexStr) = 6 Then
        Dim r As Long, g As Long, b As Long
        r = CLng("&H" & Mid$(hexStr, 1, 2))
        g = CLng("&H" & Mid$(hexStr, 3, 2))
        b = CLng("&H" & Mid$(hexStr, 5, 2))
        HexToLong = r + 256& * g + 65536 * b    'BGR order required by Excel
    Else
        HexToLong = xlNone
    End If
End Function


'---------------------------------------------------------------
'  OUTPUT  –  SUMMARY  (with colour swatches)
'---------------------------------------------------------------
Private Sub DumpSummary(sum As Object)
    
    ReplaceSheet "Color Summary"
    
    Dim sht As Worksheet: Set sht = Worksheets("Color Summary")
    sht.Range("A1:B1").Value = Array("Color (hex)", "Occurrences")
    
    Dim r As Long: r = 2
    Dim k As Variant
    For Each k In sum.Keys
        sht.Cells(r, 1).Value = k
        sht.Cells(r, 2).Value = sum(k)
        
        'Add a colour swatch
        If Left$(k, 1) = "#" Then
            sht.Cells(r, 1).Interior.Color = HexToLong(k)
        End If
        r = r + 1
    Next k
    
    sht.Columns.AutoFit
End Sub


'---------------------------------------------------------------
'  OUTPUT  –  DETAIL  (with colour swatches)
'---------------------------------------------------------------
Private Sub DumpDetail(det As Object)
    
    ReplaceSheet "Color Audit"
    
    Dim sht As Worksheet: Set sht = Worksheets("Color Audit")
    sht.Range("A1:D1").Value = _
        Array("Color (hex)", "Object Type", "Location", "Occurrences")
    
    Dim r As Long: r = 2
    Dim k As Variant, parts As Variant
    For Each k In det.Keys
        parts = Split(k, "|")                 'Color|Type|Loc
        sht.Cells(r, 1).Resize(1, 3).Value = parts
        sht.Cells(r, 4).Value = det(k)
        
        'Add a colour swatch
        If Left$(parts(0), 1) = "#" Then _
            sht.Cells(r, 1).Interior.Color = HexToLong(parts(0))
        
        r = r + 1
    Next k
    
    sht.Columns.AutoFit
End Sub


'---------------------------------------------------------------
'  UTILITY – DELETE & RECREATE REPORT SHEET
'---------------------------------------------------------------
Private Sub ReplaceSheet(name As String)
    Application.DisplayAlerts = False
    On Error Resume Next: Worksheets(name).Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    Worksheets.Add(Before:=Sheets(1)).name = name
End Sub