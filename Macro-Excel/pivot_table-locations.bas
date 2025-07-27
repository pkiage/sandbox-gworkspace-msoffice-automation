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

Sub ListPivotTableLocations()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim summaryWS As Worksheet
    Dim row As Long

    ' Create new summary sheet
    Set summaryWS = ThisWorkbook.Sheets.Add
    summaryWS.Name = "PivotTable Locations"
    summaryWS.Cells(1, 1).Value = "Worksheet"
    summaryWS.Cells(1, 2).Value = "PivotTable Name"
    summaryWS.Cells(1, 3).Value = "Top Left Cell"
    summaryWS.Cells(1, 4).Value = "Full Range Address"

    row = 2
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            summaryWS.Cells(row, 1).Value = ws.Name
            summaryWS.Cells(row, 2).Value = pt.Name
            summaryWS.Cells(row, 3).Value = pt.TableRange2.Cells(1, 1).Address
            summaryWS.Cells(row, 4).Value = pt.TableRange2.Address
            row = row + 1
        Next pt
    Next ws

    MsgBox "Pivot table locations listed."
End Sub
