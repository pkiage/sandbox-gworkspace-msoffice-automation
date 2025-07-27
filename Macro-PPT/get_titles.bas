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

Sub ExtractAllTitlesWithFallback()
    Dim sld As Slide
    Dim shp As Shape
    Dim output As String
    Dim candidate As Shape
    Dim maxFontSize As Single
    Dim thisFontSize As Single
    Dim txtRange As TextRange
    Dim iRun As Long
    
    output = "Slide Titles:" & vbCrLf & vbCrLf
    
    For Each sld In ActivePresentation.Slides
        Dim titleText As String
        
        ' 1) If there's a real Title placeholder, use it
        If sld.Shapes.HasTitle Then
            titleText = Trim(sld.Shapes.Title.TextFrame.TextRange.Text)
        
        ' 2) Else scan every text shape for the one with the largest font
        Else
            Set candidate = Nothing
            maxFontSize = 0
            
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        Set txtRange = shp.TextFrame.TextRange
                        
                        ' Find the largest font size in this shape
                        For iRun = 1 To txtRange.Runs.Count
                            thisFontSize = txtRange.Runs(iRun).Font.Size
                            If thisFontSize > maxFontSize Then
                                maxFontSize = thisFontSize
                                Set candidate = shp
                            ' If equal size, pick the shape nearest the top
                            ElseIf thisFontSize = maxFontSize Then
                                If Not candidate Is Nothing Then
                                    If shp.Top < candidate.Top Then
                                        Set candidate = shp
                                    End If
                                End If
                            End If
                        Next iRun
                    End If
                End If
            Next shp
            
            If Not candidate Is Nothing Then
                titleText = Trim(candidate.TextFrame.TextRange.Text)
            Else
                titleText = "[No Title]"
            End If
        End If
        
        output = output & "Slide " & sld.SlideIndex & ": " & titleText & vbCrLf
    Next sld
    
    ' 3) Dump the results onto a new slide
    Dim summary As Slide
    Set summary = ActivePresentation.Slides.Add( _
        ActivePresentation.Slides.Count + 1, ppLayoutText)
    summary.Shapes(1).TextFrame.TextRange.Text = "All Slide Titles"
    summary.Shapes(2).TextFrame.TextRange.Text = output
    
    MsgBox "Done! A summary slide has been added.", vbInformation
End Sub
