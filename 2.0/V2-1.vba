Sub Main()
   
      On Error GoTo ErrorHandler
        Call TimeFormat
        Call Latestorno
        Call TableFiltering
  
    
    Exit Sub
 
ErrorHandler:
        MsgBox "Error # " & Err.Number & ": " & Err.Description
        Resume Next
        End Sub

Private Sub TimeFormat()
    'Formulates the date time in a more excel friendly way and also retaining the original dates
    'Also tells if an entry should go into current, recent or future outages
    Application.StatusBar = "Running TimeFormat function... "
    'Code Under statusbar

        'Insert new columns for new datetime column data
            Columns("H:H").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Columns("J:J").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Columns("L:L").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        'Apply Formula to first data cell of each column - Pub date, start date and end date
            Range("H2").Select
            ActiveCell.FormulaR1C1 = _
                "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
            Range("J2").Select
            ActiveCell.FormulaR1C1 = _
                "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
            Range("L2").Select
            ActiveCell.FormulaR1C1 = _
                "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
    
        'Assign new names to the new columns and hide the older one's
            Range("H1").Select
            ActiveCell.FormulaR1C1 = "Publication Date"
            Range("J1").Select
            ActiveCell.FormulaR1C1 = "Start Date"
            Range("L1").Select
            ActiveCell.FormulaR1C1 = "End Date"
            Columns("G:G").Select
            Selection.EntireColumn.Hidden = True
            Columns("I:I").Select
            Selection.EntireColumn.Hidden = True
            Columns("K:K").Select
            Selection.EntireColumn.Hidden = True
        
        'Apply the formulas to all cells below them using nested function
        Call ApplyDateFormatFormula
    
    
    Application.StatusBar = False

    
   
    
    End Sub
 
Private Sub ApplyDateFormatFormula()
    'This function is a sub function for the TimeFormat function to facilitate applying the same formula for each row in all date columns and version verification column
        Dim ws1 As Worksheet
        Dim ipub As Long
        Dim isd As Long
        Dim ied As Long
        Dim lastA As Long
        

        Set ws = ThisWorkbook.ActiveSheet
        lastA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row 'Starts at row 2 and goes to the last non-empty row in column A
        For ipub = 2 To lastA 
            ws.Cells(ipub, "H").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
        Next ipub
        For isd = 2 To lastA 
            ws.Cells(isd, "J").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
        Next isd
        For ied = 2 To lastA 
            ws.Cells(ied, "L").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))"
        Next ied
     End Sub

Private Sub Latestorno()
    'First sorts ID as A to Z, then sorts version as smallest to largest, then simply uses boolean to compared ID's and tell if version is latest
       
        'Sorting
            With ActiveWorkbook.ActiveSheet
            .Range("A:Q").Select
            .Sort.SortFields.Clear
            .Sort.SortFields.Add Key:=Range("A2:A" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Sort.SortFields.Add Key:=Range("F2:F" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Sort.SetRange Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row)
            .Sort.Header = xlYes
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.SortMethod = xlPinYin
            .Sort.Apply
            End With
        'Renaming and Latest version check formula, 1 is latest, 0 is not latest    
            Range("Q1").Select
            ActiveCell.FormulaR1C1 = "Latest"
                Range("Q2").Select
                ActiveCell.FormulaR1C1 = "=IF(RC[-16]=R[1]C[-16],""0"",""1"")"
                Dim ws1 As Worksheet
                Dim iver As Long
                Dim lastA As Long
                

                Set ws1 = ThisWorkbook.ActiveSheet
                lastA = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).row 'Starts at row 2 and goes to the last non-empty row in column A
                For iver = 2 To lastA 
                    ws1.Cells(iver, "Q").Formula = "=IF(RC[-16]=R[1]C[-16],""0"",""1"")"
                Next iver
        End Sub

Private Sub TableFiltering
    'Filters with basic POWER/FR criteria, >=800 MW and 0MW avail etc.
    'Also fixes the latest version column by pasting as values and deleting the older one
        Columns("Q:Q").Select
        Selection.Copy
        Columns("R:R").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Columns("Q:Q").Select
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlToLeft

    'Filtering
       
        Columns("A:Q").Select
        Selection.AutoFilter
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=16, Criteria1:="=0", _
            Operator:=xlOr, Criteria2:="="
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=2, Criteria1:="Active"
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:= _
            "=Fortuite", Operator:=xlOr, Criteria2:="=Planifiée"
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:= _
            "=Fortuite", Operator:=xlOr, Criteria2:="=Planifiée"
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=4, Criteria1:= _
            "Nucléaire"
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=5, Criteria1:= _
        "<>*FESSENHEIM*", Operator:=xlAnd
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=15, Criteria1:=">=800" _
            , Operator:=xlAnd
        ActiveSheet.Range("A1:Q" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=17, Criteria1:="1"
End Sub