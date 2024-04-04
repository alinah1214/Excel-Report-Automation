Attribute VB_Name = "Module1"
Option Explicit

Sub Rectifier_Module()

Dim Fullpath As String, c As Integer, i As Integer, count As Integer

Application.ScreenUpdating = "False"

Fullpath = Application.ActiveWorkbook.Path
Workbooks.add.SaveAs Fullpath & "\Nokia Land.xlsx"

Workbooks.Open Fullpath & "\NETACT1.csv"
Workbooks.Open Fullpath & "\NETACT2.csv"
Workbooks("NETACT1.csv").Activate
Range("A1").Select
Selection.CurrentRegion.Select
Selection.Copy
Workbooks("Nokia Land.xlsx").Sheets(1).Paste
c = Range("A1").End(xlDown).Row

Workbooks("NETACT2.csv").Activate


Range("A1").CurrentRegion.Select

Selection.Copy
Workbooks("Nokia Land.xlsx").Activate
Range("A" & c + 1).Select
ActiveSheet.Paste


Rows(c + 1).Delete
Columns("A:C").Delete
Columns("C:O").Delete
Columns("D:H").Delete
Columns("E:U").Delete

''-------------------------Filter
filter_Sites
''---------------------------------Adding 6 Digit
Digits
''-----------------------------Region Lookup
Region_LK


''--------------------------Aging
Aging
''----------------------------------Pivot Table
Pivot
''----------------------------------Format
Format
''-------------------------------------------

Workbooks("NETACT1.csv").Close SaveChanges:=False
Workbooks("NETACT2.csv").Close SaveChanges:=False
Workbooks("Nokia Land.xlsx").Close SaveChanges:=True

''------------------------------ZTE----------------


Workbooks.add.SaveAs Fullpath & "\ZTE Land.xlsx"

Workbooks.Open Fullpath & "\EMS1.csv"
Workbooks.Open Fullpath & "\EMS2.csv"
Workbooks.Open Fullpath & "\EMS3.csv"

Workbooks("EMS1.csv").Activate
Rows(1).Delete
Rows(Range("A1").End(xlDown).Row).Delete

''''
Range("A1").CurrentRegion.Select
Selection.Copy
Workbooks("ZTE Land.xlsx").Sheets(1).Paste
c = Range("A1").End(xlDown).Row


Workbooks("EMS2.csv").Activate
Rows(1).Delete
Rows(1).Delete
Rows(Range("A1").End(xlDown).Row).Delete
Range("A1").CurrentRegion.Select
Selection.Copy
Workbooks("ZTE Land.xlsx").Activate
Range("A" & c + 1).Select
ActiveSheet.Paste

c = Range("A1").End(xlDown).Row

Workbooks("EMS3.csv").Activate
Rows(1).Delete
Rows(1).Delete
Rows(Range("A1").End(xlDown).Row).Delete
Range("A1").CurrentRegion.Select
Selection.Copy
Workbooks("ZTE Land.xlsx").Activate
Range("A" & c + 1).Select
ActiveSheet.Paste

Columns("D").Delete
Columns("E:F").Delete
''-------------------------Filter
filter_Sites

''---------------------------------Adding 6 Digit
Digits
''-----------------------------Region Loopup
Region_LK

''--------------------------Aging
Aging
''----------------------------------Pivot Table
Pivot
''----------------------------------Format
Format
''-----------------------------------------

Workbooks("EMS1.csv").Close SaveChanges:=False
Workbooks("EMS2.csv").Close SaveChanges:=False
Workbooks("EMS3.csv").Close SaveChanges:=False
Workbooks("ZTE Land.xlsx").Close SaveChanges:=True

End Sub


