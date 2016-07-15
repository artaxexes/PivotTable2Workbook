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

If Not (objFile Is Nothing OR objFolder Is Nothing) Then
    Call PivotTable2Workbook(objFile.Self.path, objFolder.Self.path)
Else
    MsgBox "Attention you must have, my young padawan"
End If



Private Sub PivotTable2Workbook(xlsPath, folderPath)
    
	Dim thisYear
	Dim thisMonth
    
	thisYear = CStr(Year(Now()))
	thisMonth = Month(Now())

End Sub



' Show error details
Private Sub ShowErr

    MsgBox "Error: " & Err.Number & vbCrLf & "Error (Hex): " & Hex(Err.Number) & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Err.Clear

End Sub



' Receives an integer and returns the name of a valid month or "unknown"
Private Function CheckMonth(ByVal monthTarget)

    Dim strMonth
    If monthTarget >= 1 AND monthTarget <= 12 Then
        strMonth = MonthName(monthTarget)
    Else
        strMonth = "Unknown"
    End If
    CheckMonth = UCase(Left(strMonth, 1)) & LCase(Right(strMonth, Len(strMonth) - 1))

End Function



' Receives a file path for check if exists and returns a boolean
Private Function FileExists(ByVal filespec)

   Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
   FileExists = fso.FileExists(filespec)
   Set fso = Nothing

End Function



' Receives a file path for delete the file
Private Function FileDelete(ByVal filespec)

    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile(filespec)
    Set fso = Nothing

End Function