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
'------------------------------------------------------------
'  Title Audit — finds every unique title layout and shows
'  which slides use it.  Optional: also flags slides that
'  differ from the first title (“reference”) for quick fixes.
'------------------------------------------------------------

Public Sub AuditTitleConsistency_V2()
    Const SHOW_OLD_DETAIL As Boolean = False   '? set TRUE if you also want the classic “mismatch vs reference” slide

    'Grab the first title we can find so the deck isn’t empty
    Dim sld As slide, ttl As shape, refFound As Boolean
    For Each sld In ActivePresentation.Slides
        Set ttl = GetTitleShape(sld)
        If Not ttl Is Nothing Then refFound = True: Exit For
    Next
    If Not refFound Then
        MsgBox "No slide titles detected – audit cancelled.", vbExclamation
        Exit Sub
    End If
    
    'Walk every slide, fingerprint the title layout, and group
    Dim groups   As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim report() As String: ReDim report(1 To ActivePresentation.Slides.Count, 1 To 5)
    Dim rows     As Long
    Dim key      As String, refT As shape
    
    For Each sld In ActivePresentation.Slides
        Set ttl = GetTitleShape(sld)
        If ttl Is Nothing Then GoTo NextSlide
        
        '-- layout fingerprint for grouping
        key = LayoutKey(ttl)
        If groups.exists(key) Then
            groups(key) = groups(key) & ", " & sld.slideIndex
        Else
            groups.Add key, CStr(sld.slideIndex)
        End If
        
        'Detail mismatches vs *first* title in deck
        If SHOW_OLD_DETAIL Then
            If refT Is Nothing Then Set refT = ttl          'store reference once
            Dim issues$
            If ttl.Left <> refT.Left Or ttl.Top <> refT.Top Then issues = issues & "Position; "
            If ttl.Width <> refT.Width Or ttl.Height <> refT.Height Then issues = issues & "Size; "
            With ttl.TextFrame.textRange.font
                If .Name <> refT.TextFrame.textRange.font.Name Then issues = issues & "Font name; "
                If .Size <> refT.TextFrame.textRange.font.Size Then issues = issues & "Font size; "
            End With
            If Len(issues) Then
                rows = rows + 1
                report(rows, 1) = sld.slideIndex
                report(rows, 2) = Left$(issues, Len(issues) - 2)             'trim "; "
                report(rows, 3) = ttl.Left & " × " & ttl.Top
                report(rows, 4) = ttl.Width & " × " & ttl.Height
                report(rows, 5) = ttl.TextFrame.textRange.font.Name & " " & ttl.TextFrame.textRange.font.Size
                HighlightProblem ttl
            End If
        End If
NextSlide:
    Next
    
    'Build output slides
    If SHOW_OLD_DETAIL And rows > 0 Then _
        BuildReportSlide report, rows, Array("Slide #", "Mismatch type(s)", _
                                             "Left × Top", "Width × Height", _
                                             "Font (name & pt)")
                                             
    BuildLayoutGroupsSlide groups
    
    MsgBox "Audit complete – see the “Title Layout Groups” slide" & _
           IIf(SHOW_OLD_DETAIL, " (and the mismatch slide)", "") & ".", vbInformation
End Sub

'------------------------------------------------------------
'  Locate a title on a slide: placeholder, Shapes.Title, or
'  fallback (largest font text box, upper-most on ties)
'------------------------------------------------------------
Function GetTitleShape(sld As slide) As shape
    'A. Built-in shortcut
    On Error Resume Next
    Set GetTitleShape = sld.Shapes.TITLE
    On Error GoTo 0
    If Not GetTitleShape Is Nothing Then Exit Function
    
    'B. Placeholder types
    Dim shp As shape, pType As Long
    For Each shp In FlattenShapes(sld.Shapes)
        If shp.Type = msoPlaceholder Then
            pType = shp.PlaceholderFormat.Type
            If pType = ppPlaceholderTitle _
               Or pType = ppPlaceholderCenterTitle _
               Or pType = ppPlaceholderVerticalTitle Then
                Set GetTitleShape = shp: Exit Function
            End If
        End If
    Next
    
    'C. Fallback – largest font, upper-most if tie
    Dim bestSize As Single, bestTop As Single: bestTop = 1E+30
    For Each shp In FlattenShapes(sld.Shapes)
        If shp.HasTextFrame And shp.TextFrame.HasText Then
            Dim fSize As Single: fSize = MaxFontSize(shp)
            If fSize > bestSize _
               Or (fSize = bestSize And shp.Top < bestTop) Then
                bestSize = fSize: bestTop = shp.Top
                Set GetTitleShape = shp
            End If
        End If
    Next
End Function

'------------------------------------------------------------
'  Max font size used inside a shape (handles mixed formats)
'------------------------------------------------------------
Function MaxFontSize(shp As shape) As Single
    Dim tr As textRange, i As Long, mSize As Single
    Set tr = shp.TextFrame.textRange
    If tr.Runs.Count = 0 Then
        MaxFontSize = tr.font.Size
    Else
        For i = 1 To tr.Runs.Count
            If tr.Runs(i).font.Size > mSize Then mSize = tr.Runs(i).font.Size
        Next i
        MaxFontSize = mSize
    End If
End Function

'------------------------------------------------------------
'  Flatten grouped shapes into a single collection
'------------------------------------------------------------
Function FlattenShapes(shps As Shapes) As Collection
    Dim all As New Collection, shp As shape
    For Each shp In shps
        If shp.Type = msoGroup Then
            Dim subShp As shape
            For Each subShp In shp.GroupItems
                all.Add subShp
            Next
        Else
            all.Add shp
        End If
    Next
    Set FlattenShapes = all
End Function

'------------------------------------------------------------
'  Give an offending title a red dashed outline
'------------------------------------------------------------
Sub HighlightProblem(shp As shape)
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Weight = 2
        .DashStyle = msoLineDash
    End With
End Sub

'------------------------------------------------------------
'  Classic detail slide (only if SHOW_OLD_DETAIL = True)
'------------------------------------------------------------
Sub BuildReportSlide(dataArr, rows As Long, headers As Variant)
    Dim rpt As slide, tbl As shape, r As Long, c As Long, cols As Long
    Set rpt = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    rpt.Name = "Title Audit – Details"
    
    cols = UBound(headers) + 1
    Set tbl = rpt.Shapes.AddTable(rows + 1, cols, 20, 60, 900, 10 + (rows + 1) * 20)
    
    For c = 1 To cols
        tbl.Table.Cell(1, c).shape.TextFrame.textRange.Text = headers(c - 1)
        tbl.Table.Cell(1, c).shape.TextFrame.textRange.font.Bold = msoTrue
    Next
    For r = 1 To rows: For c = 1 To cols
        tbl.Table.Cell(r + 1, c).shape.TextFrame.textRange.Text = dataArr(r, c)
    Next c, r
End Sub

'------------------------------------------------------------
'  Fingerprint a title’s layout so identical ones can match
'------------------------------------------------------------
Function LayoutKey(ttl As shape) As String
    Const DP As Integer = 1    'round to 0.1 pt / pixel
    LayoutKey = _
        Format(Round(ttl.Left, DP), "0." & String(DP, "0")) & "|" & _
        Format(Round(ttl.Top, DP), "0." & String(DP, "0")) & "|" & _
        Format(Round(ttl.Width, DP), "0." & String(DP, "0")) & "|" & _
        Format(Round(ttl.Height, DP), "0." & String(DP, "0")) & "|" & _
        ttl.TextFrame.textRange.font.Name & "|" & _
        Format(Round(ttl.TextFrame.textRange.font.Size, 1), "0.0")
End Function

'------------------------------------------------------------
'  Slide listing every unique layout + its member slides
'------------------------------------------------------------
Sub BuildLayoutGroupsSlide(groups As Object)
    Dim rpt As slide, tbl As shape
    Dim parts, cols&, r&, k As Variant          '? k is Variant (valid for For Each)

    Set rpt = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    rpt.Name = "Title Layout Groups"
    
    cols = 4
    Set tbl = rpt.Shapes.AddTable(groups.Count + 1, cols, 20, 60, 900, 10 + (groups.Count + 1) * 20)
    
    '-- headers ---------------------------------------------
    tbl.Table.Cell(1, 1).shape.TextFrame.textRange.Text = "Slides"
    tbl.Table.Cell(1, 2).shape.TextFrame.textRange.Text = "Left × Top"
    tbl.Table.Cell(1, 3).shape.TextFrame.textRange.Text = "Width × Height"
    tbl.Table.Cell(1, 4).shape.TextFrame.textRange.Text = "Font (name & pt)"
    Dim c&
    For c = 1 To cols
        tbl.Table.Cell(1, c).shape.TextFrame.textRange.font.Bold = msoTrue
    Next
    
    '-- data rows -------------------------------------------
    r = 1
    For Each k In groups.Keys                   '? k is Variant, so no compile error
        parts = Split(k, "|")                   'decode fingerprint
        tbl.Table.Cell(r + 1, 1).shape.TextFrame.textRange.Text = groups(k)
        tbl.Table.Cell(r + 1, 2).shape.TextFrame.textRange.Text = parts(0) & " × " & parts(1)
        tbl.Table.Cell(r + 1, 3).shape.TextFrame.textRange.Text = parts(2) & " × " & parts(3)
        tbl.Table.Cell(r + 1, 4).shape.TextFrame.textRange.Text = parts(4) & " " & parts(5)
        r = r + 1
    Next
End Sub