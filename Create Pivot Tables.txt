Sub RunEverything()
ThisWorkbook.RefreshAll
InsertMaterialPivotTable
InsertBOQPivotTable
InsertMISCPivotTable
End Sub

Sub InsertMaterialPivotTable()

Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long

'Declare Variables
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTableMaterial").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTableMaterial"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTableMaterial")
Set DSheet = Worksheets("Project Materials")

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="MaterialPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="MaterialPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("MaterialPivotTable").PivotFields("SPN")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("MaterialPivotTable").PivotFields("Description")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
'With ActiveSheet.PivotTables("MaterialPivotTable").PivotFields("Pivot")
'.Orientation = xlColumnField
'.Position = 1
'End With

'Insert Data Field
With ActiveSheet.PivotTables("MaterialPivotTable").PivotFields("Pivot")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
.Name = "Rating"
End With

'Sort Pivot Table by Rating
ActiveSheet.PivotTables("MaterialPivotTable").PivotFields("SPN").AutoSort Order:=xlDescending, Field:="Rating"

'Format Pivot Table
ActiveSheet.PivotTables("MaterialPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MaterialPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Sub InsertBOQPivotTable()

Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long

'Declare Variables
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTableBOQ").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTableBOQ"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTableBOQ")
Set DSheet = Worksheets("Project BOQ")

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="BOQPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="BOQPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("BOQPivotTable").PivotFields("BOQ")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("BOQPivotTable").PivotFields("Description")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
'With ActiveSheet.PivotTables("BOQPivotTable").PivotFields("Pivot")
'.Orientation = xlColumnField
'.Position = 1
'End With

'Insert Data Field
With ActiveSheet.PivotTables("BOQPivotTable").PivotFields("Pivot")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
.Name = "Rating"
End With

'Sort Pivot Table by Rating
ActiveSheet.PivotTables("BOQPivotTable").PivotFields("BOQ").AutoSort Order:=xlDescending, Field:="Rating"

'Format Pivot Table
ActiveSheet.PivotTables("BOQPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("BOQPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Sub InsertMISCPivotTable()

Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long

'Declare Variables
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTableMISC").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTableMISC"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTableMISC")
Set DSheet = Worksheets("Project MISC")

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="MISCPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="MISCPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("MISCPivotTable").PivotFields("MISC")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("MISCPivotTable").PivotFields("Description")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
'With ActiveSheet.PivotTables("MISCPivotTable").PivotFields("Pivot")
'.Orientation = xlColumnField
'.Position = 1
'End With

'Insert Data Field
With ActiveSheet.PivotTables("MISCPivotTable").PivotFields("Pivot")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
.NumberFormat = "#,##0"
.Name = "Rating"
End With

'Sort Pivot Table by Rating
ActiveSheet.PivotTables("MISCPivotTable").PivotFields("MISC").AutoSort Order:=xlDescending, Field:="Rating"

'Format Pivot Table
ActiveSheet.PivotTables("MISCPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MISCPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


