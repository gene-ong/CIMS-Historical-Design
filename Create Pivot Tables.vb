Sub RunEverything()
ThisWorkbook.RefreshAll
InsertMaterialPivotTable
InsertBOQPivotTable
InsertMISCPivotTable
End Sub

Sub RunEverything2()
ThisWorkbook.RefreshAll
InsertMaterialPivotTable2
InsertBOQPivotTable2
InsertMISCPivotTable2
End Sub

Private Sub InsertMaterialPivotTable()

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

'Activate Sheet
Sheets("Project Materials").Activate
 ActiveSheet.Cells(1, 1).Select

'Remove Filters if there are any
If (Sheets("Project Materials").AutoFilterMode And Sheets("Project Materials").FilterMode) Or Sheets("Project Materials").FilterMode Then
    Sheets("Project Materials").ShowAllData
End If

'Set the Pivot column to 1
Sheets("Project Materials").Range(Cells(2, 7), Cells(Sheets("Project Materials").ListObjects("Project_Materials").Range.Rows.Count, 7)).Value = 1

'Activate Sheet
Sheets("Start here").Activate

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

'Set Sheet Tab Colour
ActiveSheet.Tab.Color = RGB(25, 25, 25)

'Format Pivot Table
ActiveSheet.PivotTables("MaterialPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MaterialPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Private Sub InsertBOQPivotTable()

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

'Activate Sheet
Sheets("Project BOQ").Activate
ActiveSheet.Cells(1, 1).Select

'Remove Filters if there are any
If (Sheets("Project BOQ").AutoFilterMode And Sheets("Project BOQ").FilterMode) Or Sheets("Project BOQ").FilterMode Then
    Sheets("Project BOQ").ShowAllData
End If

'Set the Pivot column to 1
Sheets("Project BOQ").Range(Cells(2, 4), Cells(Sheets("Project BOQ").ListObjects("Project_BOQ").Range.Rows.Count, 4)).Value = 1

'Activate Sheet
Sheets("Start here").Activate

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

'Set Sheet Tab Colour
ActiveSheet.Tab.Color = RGB(25, 25, 25)

'Format Pivot Table
ActiveSheet.PivotTables("BOQPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("BOQPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Private Sub InsertMISCPivotTable()

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

'Activate Sheet
Sheets("Project MISC").Activate
ActiveSheet.Cells(1, 1).Select

'Remove Filters if there are any
If (Sheets("Project MISC").AutoFilterMode And Sheets("Project MISC").FilterMode) Or Sheets("Project MISC").FilterMode Then
    Sheets("Project MISC").ShowAllData
End If

'Set the Pivot column to 1
Sheets("Project MISC").Range(Cells(2, 11), Cells(Sheets("Project MISC").ListObjects("Project_MISC").Range.Rows.Count, 11)).Value = 1

'Activate Sheet
Sheets("Start here").Activate

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

'Set Sheet Tab Colour
ActiveSheet.Tab.Color = RGB(25, 25, 25)

'Format Pivot Table
ActiveSheet.PivotTables("MISCPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MISCPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Private Sub InsertMaterialPivotTable2()

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

'Activate Sheet
Sheets("Project Materials (2)").Activate
ActiveSheet.Cells(1, 1).Select

'Remove Filters if there are any
If (Sheets("Project Materials (2)").AutoFilterMode And Sheets("Project Materials (2)").FilterMode) Or Sheets("Project Materials (2)").FilterMode Then
    Sheets("Project Materials (2)").ShowAllData
End If

'Set the Pivot column to 1
Sheets("Project Materials (2)").Range(Cells(2, 7), Cells(Sheets("Project Materials (2)").ListObjects("Project_Materials__2").Range.Rows.Count, 7)).Value = 1

'Activate Sheet
Sheets("Start here").Activate

Worksheets("PivotTableMaterial").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTableMaterial"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTableMaterial")
Set DSheet = Worksheets("Project Materials (2)")

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

'Set Sheet Tab Colour
ActiveSheet.Tab.Color = RGB(25, 25, 25)

'Format Pivot Table
ActiveSheet.PivotTables("MaterialPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MaterialPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Private Sub InsertBOQPivotTable2()

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

'Activate Sheet
Sheets("Project BOQ (2)").Activate
ActiveSheet.Cells(1, 1).Select

'Remove Filters if there are any
If (Sheets("Project BOQ (2)").AutoFilterMode And Sheets("Project BOQ (2)").FilterMode) Or Sheets("Project BOQ (2)").FilterMode Then
    Sheets("Project BOQ (2)").ShowAllData
End If

'Set the Pivot column to 1
Sheets("Project BOQ (2)").Range(Cells(2, 4), Cells(Sheets("Project BOQ (2)").ListObjects("Project_BOQ__2").Range.Rows.Count, 4)).Value = 1

'Activate Sheet
Sheets("Start here").Activate

Worksheets("PivotTableBOQ").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTableBOQ"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTableBOQ")
Set DSheet = Worksheets("Project BOQ (2)")

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

'Set Sheet Tab Colour
ActiveSheet.Tab.Color = RGB(25, 25, 25)

'Format Pivot Table
ActiveSheet.PivotTables("BOQPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("BOQPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Private Sub InsertMISCPivotTable2()

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

'Activate Sheet
Sheets("Project MISC (2)").Activate
ActiveSheet.Cells(1, 1).Select

'Remove Filters if there are any
If (Sheets("Project MISC (2)").AutoFilterMode And Sheets("Project MISC (2)").FilterMode) Or Sheets("Project MISC (2)").FilterMode Then
    Sheets("Project MISC (2)").ShowAllData
End If

'Set the Pivot column to 1
Sheets("Project MISC (2)").Range(Cells(2, 11), Cells(Sheets("Project MISC (2)").ListObjects("Project_MISC__2").Range.Rows.Count, 11)).Value = 1

'Activate Sheet
Sheets("Start here").Activate

Worksheets("PivotTableMISC").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTableMISC"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTableMISC")
Set DSheet = Worksheets("Project MISC (2)")

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

'Set Sheet Tab Colour
ActiveSheet.Tab.Color = RGB(25, 25, 25)

'Format Pivot Table
ActiveSheet.PivotTables("MISCPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MISCPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub



