Sub main_macro()
    ' This is the main macro which will call several nested macros...
        ' ...for easy code handling, we have nested it like this
        Columns("A:A").Select
        ActiveSheet.Paste
        Call Text_To_Column
        
        
        Call pre_max_version_clean
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
' the max version code takes high processing power as it analyzes each versio for each ID
    ' this function sorts and filters out older versions of the same REMIT message, ultimately showing us only the latest version for each REMIT ID
    ' this function also filters the table to what we need: Takes planned, nuclear and >=800 MW fully out (0MW) outages
    ' sorts date from Z to A
    ' only active outages
    Dim ws As Worksheet
    Dim rng As Range
    Dim row As Range
    Dim max_version As Variant
    Dim curr_id As Variant
    Dim last_row As Long

    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("A2:B999999")

    ' Getting the last row in the range
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
    ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=3, Criteria1:= _
            "=Fortuite", Operator:=xlOr, Criteria2:="=Planifiée"
    ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=4, Criteria1:= _
            "Nucléaire"
    ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=14, Criteria1:=">=800" _
            , Operator:=xlAnd
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=15, Criteria1:="=0", Operator:=xlOr, Criteria2:="="
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=6, Criteria1:="Active"
        Columns("K:K").ColumnWidth = 20
        Columns("I:I").ColumnWidth = 20
        
End Sub

Sub TableFilters()
' Filters table
    ' Takes planned, nuclear and >=800 MW fully out (0MW) outages
    ' sorts date from Z to A
    ' only active outages
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
        ActiveSheet.Range("$A$1:$P$999999").AutoFilter Field:=6, Criteria1:="Active"
        Columns("K:K").ColumnWidth = 20
        Columns("I:I").ColumnWidth = 20
        
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
    For ione = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Starts at row 2 and goes to the last non-empty row in column H
        ws.Cells(ione, "I").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))" 
    Next ione

End Sub

Sub ApplySectionFormula()
'This function is a sub function for the date time formula function to facilitate applying the same formula for each row in the new section column
    Set ws = ThisWorkbook.ActiveSheet
    Dim ithr As Long

    For ithr = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Starts at row 2 and goes to the last non-empty row in column H
        ws.Cells(ithr, "P").Formula = "=IF(RC[-7]>NOW(),""FUTURE"",IF(AND(RC[-7]<=NOW(),RC[-5]>=NOW()),""Current"", ""Recent""))" 'Change "=A1+B1" to your desired formula
    Next ithr


End Sub

Sub ApplyEndDateFormula()
'This function is a sub function for the date time formula function to facilitate applying the same formula for each row in the end date column
    Set ws = ThisWorkbook.ActiveSheet
    Dim itwo As Long

    For itwo = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Starts at row 2 and goes to the last non-empty row in column H
        ws.Cells(itwo, "K").Formula = "=DATE(LEFT(RC[-1],4),MID(RC[-1], 6, 2),MID(RC[-1], 9, 2))+TIME(MID(RC[-1], 12, 2),MID(RC[-1], 15, 2),MID(RC[-1], 18, 2))" 'Change "=A1+B1" to your desired formula
    Next itwo
End Sub


Sub SplitTable()
'This function splits the table into three tables- current, recent and future outages
    Call NewSheets
    Call MakeRecentTable
    Dim wsthree As Worksheet
    Dim wsfour As Worksheet
    Dim wsfive As Worksheet
    Set wsthree = ThisWorkbook.Worksheets("Current") 'Change "Sheet1" to the name of your sheet
    Set wsfour = Worksheets.Add
    Set wsfive = Worksheets.Add
    wsfour.name = "Future"
    wsfive.name = "Recent"


    Dim nowDate As Date
    nowDate = Datevalue(Now())

    Dim tbl As ListObject
    Set tbl = wsthree.ListObjects.Add(xlSrcRange, wsthree.Range("$A:$R"), , xlYes)
    tbl.Name = "Current" 
    Dim headerss As Variant
    headerss = Array("id", "version", "type", "type of generation unit", "generation unit", "Status", _
                    "publication date and time", "start date and time", "end date and time", "cause", _
                    "complements", "maximum output (MW)", "available output (MW)", "start date and time 2", "column3","end date and time 4","Column5","Column6")

    Dim tbl1 As ListObject
    Set tbl1 = wsfour.ListObjects.Add(xlSrcRange, wsfour.Range("$A:$R").CurrentRegion, , xlYes)
    tbl1.Name = "Future" 'Change "MyTable" to the name you want to give to the table
    tbl1.HeaderRowRange.Value = headerss

    Dim tbl2 As ListObject
    Set tbl2 = wsfive.ListObjects.Add(xlSrcRange, wsfive.Range("$A:$R").CurrentRegion, , xlYes)
    tbl2.Name = "Recent" 'Change "MyTable" to the name you want to give to the table
    tbl2.HeaderRowRange.Value = headerss

    Dim rowOne As Long
    Dim rowCount As Long
    rowCount = tbl.ListRows.Count

    For rowOne = rowCount To 1 Step -1
        If tbl.ListRows(rowOne).Range(1, 8).Value >= nowDate Then 'Change the date value to your first criteria
            tbl.ListRows(rowOne).Range.Cut Destination:=wsfour.ListObjects("Future").ListRows(wsfour.ListObjects("Future").ListRows.Count).Range 'Change "Table2" to the name of your second table
        ElseIf tbl.ListRows(rowOne).Range(1, 8).Value <= nowDate And tbl.ListRows(rowOne).Range(1, 9).Value < nowDate Then 'Change the date values to your second criteria
            tbl.ListRows(rowOne).Range.Cut Destination:=wsfive.ListObjects("Recent").ListRows(wsfive.ListObjects("Recent").ListRows.Count).Range 'Change "Table3" to the name of your third table
        End If
    Next rowOne



End Sub

Sub maketable()
' maketable Macro
    '

    '
        Columns("A:R").Select
        Application.CutCopyMode = False
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$R"), , xlYes).Name = _
            "Current"
        Columns("A:R").Select
End Sub

Sub NewSheets()

End Sub

Sub MakeRecentTable()
' IDK 
    Sheets.Add After:=wsthree
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "Id"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Availability"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "hasda"
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "hdasad"
        Columns("A:D").Select
        Application.CutCopyMode = False
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$D"), , xlYes).Name = _
            "Table1"
End Sub

Sub pre_max_version_clean()
    'This function is here to do a cleanup of the table before the max_version function runs
    'This is important because max_version funciton checks each row and having 10K+ to check is heavy on processing
    
    'Lvl1: Filtering out stuff mentioned in tablefilters() by selecting stuff we dont need and deleting by using visible only selection

    'Lvl2: Alternatively we can copy the data we need and then paste it beside the table and then delete the original columns

    'Lvl1.1:unfilter the table

            Range("E6").Select
            Selection.End(xlUp).Select
            Selection.End(xlToLeft).Select
            Selection.AutoFilter
            Range("A2").Select
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            ActiveSheet.Range(Selection).AutoFilter Field:=3, Criteria1:= _
                "Chronique"
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
            ActiveSheet.Range(Selection).AutoFilter Field:=3
            Range("A2").Select
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            ActiveSheet.Range(Selection).AutoFilter Field:=2, Criteria1:=Array( _
                "Annulée", "Inactive", "Supprimée"), Operator:=xlFilterValues
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
            ActiveSheet.Range(Selection).AutoFilter Field:=2
            Range("A2").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range("A2").Select
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            ActiveSheet.Range(Selection).AutoFilter Field:=4, Criteria1:=Array( _
                "Fil de l'eau et éclusé hydraulique", "Fuel / TAC", "Gaz fossile", _
                "Houille fossile", "Réservoir hydraulique", _
                "Station de transfert d'énergie par pompage hydraulique"), Operator:= _
                xlFilterValues
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
            ActiveSheet.Range(Selection).AutoFilter Field:=4
            Range("A2").Select
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            'make it not "0"
            ' same for other filters else they will pick up garbage
            ActiveSheet.Range(Selection).AutoFilter Field:=13, Criteria1:=Array( _
                "10", "100", "1050", "1070", "1125", "1150", "1160", "1165", "1170", "1180", "1190", _
                "1200", "1220", "1225", "1260", "1280", "1306", "1328", "1330", "150", "154", "165", "180" _
                , "203", "205", "210", "215", "216", "218", "230", "236", "240", "248", "25", "260", "262", _
                "270", "275", "276", "280", "286", "300", "314", "330", "348", "350", "360", "368", "370", _
                "373", "375", "390", "396", "400", "450", "460", "473", "480", "513", "525", "530", "535", _
                "540", "545", "558", "560", "580", "585", "590", "598", "60", "600", "630", "645", "650", _
                "660", "670", "675", "686", "690", "710", "725", "73", "730", "734", "740", "750", "755", _
                "758", "763", "770", "775", "780", "788", "790", "795", "800", "801", "805", "806", "809", _
                "83", "870", "875", "880", "900", "903", "910", "916", "920", "95", "950", "970", "980"), _
                Operator:=xlFilterValues
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
            ActiveSheet.Range("$A$1:$M$222").AutoFilter Field:=13
            Range("A2").Select
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            ActiveWindow.SmallScroll Down:=-135
            ActiveWindow.ScrollRow = 225
            ActiveWindow.ScrollRow = 223
            ActiveWindow.ScrollRow = 222
            ActiveWindow.ScrollRow = 220
            ActiveWindow.ScrollRow = 219
            ActiveWindow.ScrollRow = 214
            ActiveWindow.ScrollRow = 212
            ActiveWindow.ScrollRow = 209
            ActiveWindow.ScrollRow = 208
            ActiveWindow.ScrollRow = 205
            ActiveWindow.ScrollRow = 202
            ActiveWindow.ScrollRow = 199
            ActiveWindow.ScrollRow = 196
            ActiveWindow.ScrollRow = 191
            ActiveWindow.ScrollRow = 184
            ActiveWindow.ScrollRow = 182
            ActiveWindow.ScrollRow = 176
            ActiveWindow.ScrollRow = 173
            ActiveWindow.ScrollRow = 167
            ActiveWindow.ScrollRow = 162
            ActiveWindow.ScrollRow = 159
            ActiveWindow.ScrollRow = 151
            ActiveWindow.ScrollRow = 150
            ActiveWindow.ScrollRow = 148
            ActiveWindow.ScrollRow = 149
            ActiveWindow.ScrollRow = 150
            ActiveWindow.ScrollRow = 151
            ActiveWindow.ScrollRow = 153
            ActiveWindow.ScrollRow = 154
            ActiveWindow.ScrollRow = 155
            ActiveWindow.ScrollRow = 156
            ActiveWindow.ScrollRow = 157
            ActiveWindow.ScrollRow = 158
            ActiveWindow.ScrollRow = 159
            ActiveWindow.ScrollRow = 160
            ActiveWindow.ScrollRow = 161
            ActiveWindow.ScrollRow = 162
            ActiveWindow.ScrollRow = 163
            ActiveWindow.ScrollRow = 164
            ActiveWindow.ScrollRow = 165
            ActiveWindow.ScrollRow = 166
            ActiveWindow.ScrollRow = 167
            ActiveWindow.ScrollRow = 168
            ActiveWindow.ScrollRow = 169
            ActiveWindow.ScrollRow = 170
            ActiveWindow.ScrollRow = 171
            ActiveWindow.ScrollRow = 172
            ActiveWindow.ScrollRow = 170
            ActiveWindow.ScrollRow = 166
            ActiveWindow.ScrollRow = 159
            ActiveWindow.ScrollRow = 148
            ActiveWindow.ScrollRow = 119
            ActiveWindow.ScrollRow = 101
            ActiveWindow.ScrollRow = 80
            ActiveWindow.ScrollRow = 53
            ActiveWindow.ScrollRow = 38
            ActiveWindow.ScrollRow = 29
            ActiveWindow.ScrollRow = 20
            ActiveWindow.ScrollRow = 18
            ActiveWindow.ScrollRow = 14
            ActiveWindow.ScrollRow = 13
            ActiveWindow.ScrollRow = 12
            ActiveWindow.ScrollRow = 11
            ActiveWindow.ScrollRow = 10
            ActiveWindow.ScrollRow = 8
            ActiveWindow.ScrollRow = 6
            ActiveWindow.ScrollRow = 5
            ActiveWindow.ScrollRow = 3
            ActiveWindow.ScrollRow = 1
            Range("A2").Select
            ActiveSheet.Range("$A$1:$M$222").AutoFilter Field:=12, Criteria1:="<800", _
                Operator:=xlAnd
            ActiveSheet.Range("$A$1:$M$222").AutoFilter Field:=12
            Range("A2").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range("I4").Select
            ActiveWorkbook.Save
            Windows("Try 5.xlsm").Activate
        

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
    ' Shifitng to new columns and then deleting the older data
End Sub
    