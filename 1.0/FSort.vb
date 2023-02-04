
Sub FSort()
'
' FSort Macro
' Sorts French nuke CSV file.
'

'
    Range("H14").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:= _
        True
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "inkypinkyponky"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Start Date"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "End Date"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Sector"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Type of production unit"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Name of unit"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Maximum power"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Available power"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]=R[-1]C[-1]-1)&(RC[1]=R[-1]C[1])"
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C2:C3"), Type:=xlFillDefault
    Range("C2:C3").Select
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C10301"), Type:=xlFillDefault
    Range("C3:C10301").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$R$301").AutoFilter Field:=3, Criteria1:= _
        "=FALSEFALSE", Operator:=xlOr, Criteria2:="=#VALUE!"
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:F").ColumnWidth = 15.56
    Columns("G:G").ColumnWidth = 15.11
    Columns("H:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:K").ColumnWidth = 19.22
    ActiveSheet.Range("$A$1:$R$301").AutoFilter Field:=11, Criteria1:="Nuclear"
    Columns("N:N").Select
    Selection.EntireColumn.Hidden = True
    Columns("O:O").ColumnWidth = 15.78
    ActiveSheet.Range("$A$1:$R$301").AutoFilter Field:=15, Criteria1:=">=800", _
        Operator:=xlAnd
    ActiveSheet.Range("$A$1:$R$301").AutoFilter Field:=16, Criteria1:=""
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    
End Sub


