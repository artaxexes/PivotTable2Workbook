' File: pivotTable2Workbook.vbs
' Description: Create workbook for pivot tables based on named columns (month, product, value)
' Author/Company: Artaxexes J. Ferreira/Biggdata
' Contact: anddrei.ferreira@biggdata.com.br
' PivotTable ref.: https://msdn.microsoft.com/en-us/library/office/ff837611.aspx

Option Explicit

Const srcWkbFilepath = "path/to/source/file/with/pivotTables.xls(x)?"
Const resultFolder = "path/to/save/result/file/"

Call wbkFromPivotTable(srcWkbFilepath, resultFolder)

Private Sub wbkFromPivotTable(xlSrcWkbFilepath, resultFolder)

End Sub

Private Sub ShowErr

End Sub
