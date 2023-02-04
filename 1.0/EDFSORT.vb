Sub EDFSORT()
'
' Macro1 Macro
'

'
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        )), TrailingMinusNumbers:=True
    Columns("E:E").Select
    Selection.AutoFilter
    Range("E1").Select
    ActiveSheet.Range("$E$1:$E$10386").AutoFilter Field:=1, Criteria1:= _
        "Nucléaire"
    Columns("F:F").ColumnWidth = 13.78
    Columns("G:G").ColumnWidth = 15.11
    Columns("C:C").ColumnWidth = 5
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Version No."
    Range("C2").Select
    Columns("D:D").ColumnWidth = 11.78
    Columns("D:D").ColumnWidth = 12.56
    Columns("I:I").ColumnWidth = 11.56
    Columns("I:I").ColumnWidth = 13.67
    Columns("H:H").Select
    Selection.AutoFilter
    Columns("A:M").Select
    Selection.AutoFilter
    Range("E1").Select
    ActiveSheet.Range("$A$1:$M$10386").AutoFilter Field:=5, Criteria1:= _
        "Nucléaire"
    ActiveSheet.Range("$A$1:$M$10386").AutoFilter Field:=8, Criteria1:= _
        "=Fortuite", Operator:=xlOr, Criteria2:="=Planifiée"
    Columns("M:M").ColumnWidth = 13.44
    ActiveSheet.Range("$A$1:$M$10386").AutoFilter Field:=12, Criteria1:="=0", _
        Operator:=xlOr, Criteria2:="="
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Filter"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Start date"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "End date"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Maximum capacity (MW)"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Unavailable capacity (MW)"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Date of publication"
    Range("M3").Select
    Columns("M:M").ColumnWidth = 16.44
    ActiveWindow.SmallScroll Down:=0
    Columns("M:M").ColumnWidth = 18.89
    Columns("M:M").ColumnWidth = 17.67
    Range("M1").Select
    ActiveWorkbook.Worksheets("Export_toutes_versions (3)").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Export_toutes_versions (3)").AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("M1:M10386"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Export_toutes_versions (3)").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F5").Select
    ActiveWindow.SmallScroll Down:=0
    ActiveSheet.Range("$A$1:$M$10386").AutoFilter Field:=11, Criteria1:=">=800" _
        , Operator:=xlAnd
    Range("A3").Select
End Sub
