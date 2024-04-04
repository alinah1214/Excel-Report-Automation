Attribute VB_Name = "Module2"
Option Explicit


Sub Pivot()
Dim PTable As PivotTable
Dim PCache As PivotCache
Dim PRange As Range
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim LR As Long
Dim LC As Long



On Error Resume Next

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Worksheets("Pivot Sheet").Delete 'This will delete the exisiting pivot table worksheet
Worksheets.add After:=ActiveSheet ' This will add new worksheet
ActiveSheet.Name = "Pivot Sheet" ' This will rename the worksheet as "Pivot Sheet"
On Error GoTo 0

Set PSheet = Worksheets("Pivot Sheet")
Set DSheet = Worksheets("Sheet1")

'Find Last used row and column in data sheet
LR = DSheet.Cells(Rows.count, 1).End(xlUp).Row
LC = DSheet.Cells(1, Columns.count).End(xlToLeft).Column

'Set the pivot table data range
Set PRange = DSheet.Cells(1, 1).Resize(LR, LC)

'Set pivot cahe
Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)

'Create blank pivot table
Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="Sales_Report")

 

ActiveSheet.PivotTables("Sales_report").AddDataField ActiveSheet.PivotTables( _
        "Sales_Report").PivotFields("Region"), "Count of Region", xlCount
With PSheet.PivotTables("Sales_Report").PivotFields("Region")
.Orientation = xlRowField
.Position = 1
End With

End Sub
Sub Digits()
Dim i As Integer
Columns("D").Insert
Range("D1") = "NE"
For i = 2 To Range("A1").End(xlDown).Row
 If Left(Range("C" & i), 1) = "u" Or Left(Range("C" & i), 1) = "e" Then
  Range("D" & i) = Mid(Range("C" & i), 2, 6)
  Else
  Range("D" & i) = Left(Range("C" & i), 6)
  End If
 Next i

End Sub

Sub Aging()
Range("G1") = "Aging"
Range("G2").Select
 ActiveCell.FormulaR1C1 = "=NOW()-RC[-5]"
'  Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & Range("A2").End(xlDown).Row)
    Range("G2:G" & Range("A2").End(xlDown).Row).Select
    Selection.NumberFormat = "dd"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("G:G").Select
    Selection.NumberFormat = "0"

Range("A1:G" & Range("A1").End(xlDown).Row).AutoFilter Field:=7, Criteria1:=0, Operator:=xlOr, Criteria2:=1
Range("A2").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Range("A:A").AutoFilter
 
Range("A1:G" & Range("A1").End(xlDown).Row).AutoFilter Field:=7, Criteria1:=2, Operator:=xlOr, Criteria2:=3
Range("A2").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Range("A:A").AutoFilter

End Sub

Sub Region_LK()

Range("f1") = "Region"
    Range("f2").Select
    
ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'[Rectifier Module Faulty Macro.xlsm]Sheet1'!C1:C2,2,0)"
        Selection.AutoFill Destination:=Range("F2:F" & Range("A2").End(xlDown).Row)
     Range("F2:F" & Range("A2").End(xlDown).Row).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Range("A1:G" & Range("A1").End(xlDown).Row).AutoFilter Field:=6, Criteria1:="#N/A"
Range("A2").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Range("A:A").AutoFilter


End Sub

Sub filter_Sites()

Range("A1:D" & Range("A1").End(xlDown).Row).AutoFilter Field:=3, Criteria1:="", Operator:=xlOr, Criteria2:="*TOF*"
Range("A2").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Range("A:A").AutoFilter

Range("A1:D" & Range("A1").End(xlDown).Row).AutoFilter Field:=3, Criteria1:="*DTP*", Operator:=xlOr, Criteria2:="*LCK*"
Range("A2").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Range("A:A").AutoFilter

End Sub
Sub Format()

Sheets(1).Columns("A:G").AutoFit
Sheets(1).Range("A1:G1").Font.Bold = True
Sheets(1).Range("A1:G1").Interior.ColorIndex = 6
Sheets(1).Columns("A:G").VerticalAlignment = xlCenter
Sheets(1).Columns("A:G").HorizontalAlignment = xlCenter
Sheets(1).UsedRange.Borders.LineStyle = xlContinuous

End Sub
