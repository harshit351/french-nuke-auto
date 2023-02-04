Sub filter_table()
Dim ws As Worksheet
Dim rng As Range
Dim row As Range
Dim max_version As Variant
Dim curr_id As Variant
Dim last_row As Long

' Set the worksheet and range variables
Set ws = ThisWorkbook.Worksheets("Sheet1")
Set rng = ws.Range("A2:B10")

' Get the last row in the range
last_row = rng.Rows(rng.Rows.Count).Row

' Loop through each row in the range
For Each row In rng.Rows
    ' Get the current ID and version
    curr_id = row.Cells(1).Value
    curr_version = row.Cells(2).Value
    
    ' If the current ID is different from the previous one,
    ' reset the maximum version
    If curr_id <> prev_id Then
        max_version = curr_version
    End If
    
    ' If the current version is greater than the maximum version,
    ' update the maximum version
    If curr_version > max_version Then
        max_version = curr_version
    End If
    
    ' If the current row is the last row in the range,
    ' filter the table according to the maximum version
    If row.Row = last_row Then
        rng.AutoFilter Field:=2, Criteria1:=max_version
    End If
    
    ' Set the previous ID to the current ID
    prev_id = curr_id
Next row
End Sub
