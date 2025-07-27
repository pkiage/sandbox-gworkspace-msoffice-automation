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

Sub ShowOnlyLinkedPivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pivotInfo As Object
    Dim cacheIndex As Variant
    Dim summaryWS As Worksheet
    Dim row As Long
    Dim itemList As Variant
    Dim item As Variant

    ' Create a dictionary to group PivotTables by CacheIndex
    Set pivotInfo = CreateObject("Scripting.Dictionary")

    ' Collect all pivot tables and group by their CacheIndex
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            cacheIndex = pt.CacheIndex
            If Not pivotInfo.Exists(cacheIndex) Then
                pivotInfo.Add cacheIndex, _
                    Array(ws.Name & " - " & pt.Name & " @ " & _
                          pt.TableRange2.Cells(1, 1).Address)
            Else
                itemList = pivotInfo(cacheIndex)
                ReDim Preserve itemList(UBound(itemList) + 1)
                itemList(UBound(itemList)) = _
                    ws.Name & " - " & pt.Name & " @ " & _
                    pt.TableRange2.Cells(1, 1).Address
                pivotInfo(cacheIndex) = itemList
            End If
        Next pt
    Next ws

    ' Delete existing summary sheet (if any)
    On Error Resume Next
      Application.DisplayAlerts = False
      ThisWorkbook.Worksheets("PivotTable Links").Delete
      Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new summary sheet
    Set summaryWS = ThisWorkbook.Worksheets.Add
    summaryWS.Name = "PivotTable Links"
    summaryWS.Cells(1, 1).Value = "Cache Index"
    summaryWS.Cells(1, 2).Value = "Pivot Tables Sharing Cache"
    row = 2

    ' Output only those CacheIndices that have more than one PivotTable
    For Each cacheIndex In pivotInfo.Keys
        itemList = pivotInfo(cacheIndex)
        If UBound(itemList) >= 1 Then
            summaryWS.Cells(row, 1).Value = cacheIndex
            ' Join all entries for this cache into one cell with newline
            summaryWS.Cells(row, 2).Value = Join(itemList, vbNewLine)
            row = row + 1
        End If
    Next cacheIndex

    ' Autofit columns for readability
    summaryWS.Columns("A:B").AutoFit

    MsgBox "Only linked PivotTables have been listed.", vbInformation
End Sub