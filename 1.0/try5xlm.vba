Sub main_macro()
    ' This is the main macro which will call several nested macros...
        ' ...for easy code handling, we have nested it like this
       
        Call Text_To_Column
        Call Swap_Column
        Call DateTimeFormat
        Call max_version
        Call ShiftCleanLvl1

End Sub

Sub Text_To_Column()
' This function converts our raw csv data into excel columns for further sorting
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        )), TrailingMinusNumbers:=True
End Sub

Sub Swap_Column()
' This macro swaps the version column into column B for easier handling of the max_version function
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Copy
    Columns("G:G").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
End Sub

Sub max_version()
    Dim ws As Worksheet
    Dim rng As Range
    Dim row As Range
    Dim max_version As Variant
    Dim curr_id As Variant
    Dim last_row As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("A1:B999999")
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row - 1
    Debug.Print curr_id, curr_version, max_version
    For Each row In rng.Rows
        curr_id = row.Cells(1).Value
        curr_version = row.Cells(2).Value
        If curr_id <> prev_id Then
            max_version = curr_version
        End If
        If curr_version > max_version Then
            max_version = curr_version
        End If
        prev_id = curr_id
    Next row
    rng.AutoFilter Field:=2, Criteria1:=max_version
End Sub



Sub DateTimeFormat()
'Formulates the date time in a more excel friendly way and also retaining the original dates
    'Also tells if an entry should go into current, recent or future outages
        Columns("I:I").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns("K:K").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("I1").Select
        ActiveCell.FormulaR1C1 = "start date"
        Range("K1").Select
        ActiveCell.FormulaR1C1 = "end date"
        Range("I2").Select
        ActiveCell.FormulaR1C1 = _
            "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
        Range("K2").Select
        ActiveCell.FormulaR1C1 = _
            "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
        Range("P1").Select
        ActiveCell.FormulaR1C1 = "Section"
        Range("P2").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(RC[-7]>NOW(),""FUTURE"",IF(AND(RC[-7]<=NOW(),RC[-5]>=NOW()),""Current"", ""Recent""))"
        Columns("H:H").Select
        Selection.EntireColumn.Hidden = True
        Columns("J:J").Select
        Selection.EntireColumn.Hidden = True
        Columns("I:I").Select
        Selection.NumberFormat = "mm/dd/yyyy hh:mm:ss"
        Selection.ColumnWidth = 14.44
        Selection.ColumnWidth = 19.33
        Columns("K:K").Select
        Selection.NumberFormat = "mm/dd/yyyy hh:mm:ss"
        Selection.ColumnWidth = 17.22
    ' These subfunctions apply the formula to the whole column
        Call ApplyStartDateFormula
        Call ApplyEndDateFormula
        Call ApplySectionFormula
End Sub

Sub ApplyStartDateFormula()
'This function is a sub function for the date time formula function to facilitate applying the same formula for each row in the start date column
    Dim ws As Worksheet
    Dim ione As Long
    Set ws = ThisWorkbook.ActiveSheet
    For ione = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row 'Starts at row 2 and goes to the last non-empty row in column H
        ws.Cells(ione, "I").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
    Next ione

End Sub

Sub ApplySectionFormula()
'This function is a sub function for the date time formula function to facilitate applying the same formula for each row in the new section column
    Set ws = ThisWorkbook.ActiveSheet
    Dim ithr As Long

    For ithr = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row 'Starts at row 2 and goes to the last non-empty row in column H
        ws.Cells(ithr, "P").Formula = "=IF(RC[-7]>NOW(),""FUTURE"",IF(AND(RC[-7]<=NOW(),RC[-5]>=NOW()),""Current"", ""Recent""))" 'Change "=A1+B1" to your desired formula
    Next ithr


End Sub

Sub ApplyEndDateFormula()
'This function is a sub function for the date time formula function to facilitate applying the same formula for each row in the end date column
    Set ws = ThisWorkbook.ActiveSheet
    Dim itwo As Long

    For itwo = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row 'Starts at row 2 and goes to the last non-empty row in column H
        ws.Cells(itwo, "K").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))" 'Change "=A1+B1" to your desired formula
    Next itwo
End Sub

Sub TableFilters()
' Filters table
    ' Takes planned, nuclear and >=800 MW fully out (0MW) outages
    ' sorts date from Z to A
    Columns("A:P").Select
        Selection.AutoFilter
        Range("N1").Select
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=3, Criteria1:= _
            "=Fortuite", Operator:=xlOr, Criteria2:="=Planifiée"
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=4, Criteria1:= _
            "Nucléaire"
        Range("L1").Select
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=14, Criteria1:=">=800" _
            , Operator:=xlAnd
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=15, Criteria1:="=0", Operator:=xlOr, Criteria2:="="
        Columns("K:K").ColumnWidth = 20
        Columns("I:I").ColumnWidth = 20
End Sub

Sub ShiftCleanLvl1()
    ' Hiding non relevant POWER/FR rows
        Columns("A:A").Select
        Selection.EntireColumn.Hidden = True
        Columns("B:B").Select
        Selection.EntireColumn.Hidden = True
        Columns("D:D").Select
        Selection.EntireColumn.Hidden = True
        Columns("F:F").Select
        Selection.EntireColumn.Hidden = True
        Columns("G:G").Select
        Selection.EntireColumn.Hidden = True
        Columns("H:H").Select
        Selection.EntireColumn.Hidden = True
        Columns("J:J").Select
        Selection.EntireColumn.Hidden = True
        Columns("L:L").Select
        Selection.EntireColumn.Hidden = True
        Columns("M:M").Select
        Selection.EntireColumn.Hidden = True
End Sub
    

