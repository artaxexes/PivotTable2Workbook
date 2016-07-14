' File: pivotTable2Workbook.vbs
' Description: Create workbook for pivot tables based on named columns (month, product, value)
' Author/Company: Artaxexes J. Ferreira/Biggdata
' Contact: anddrei.ferreira@biggdata.com.br
' PivotTable ref.: https://msdn.microsoft.com/en-us/library/office/ff837611.aspx

Option Explicit



Dim objShell, objFile, objFolder

Set objShell = CreateObject("Shell.Application")
Set objFile = objShell.BrowseForFolder(0, "Select the MS Excel file with target pivot tables:", &H4000, "")
Set objFolder = objShell.BrowseForFolder(0, "Please select the folder to save result file:", 1, "")

If Not (objFile Is Nothing or objFolder Is Nothing) Then
    Call pivotTable2Workbook(objFile.Self.path, objFolder.Self.path)
Else
    MsgBox "Attention you must have, my young padawan"
End If



Private Sub pivotTable2Workbook(xlsPath, folderPath)

    MsgBox "xls path: " & xlsPath
    MsgBox "Folder to save: " & folderPath
    
	Dim thisYear
	Dim thisMonth
    
	thisYear = CStr(Year(Now()))
	thisMonth = Month(Now())
	MsgBox thisYear & ", " & thisMonth

End Sub



Private Sub ShowErr

End Sub