' File: pivotTable2Workbook.vbs
' Description: Create workbook for pivot tables based on named columns (eg., month, product, value)
' Author: Artaxexes J. Ferreira
' Contact: artaxexes@ymail.com
' PivotTable ref.: https://msdn.microsoft.com/en-us/library/office/ff837611.aspx

Option Explicit



Const resultFolder = "C:\Users\anddrei.ferreira\Documents\jobs\07_anp\Test\out\"
Const sourceFolder = "C:\Users\anddrei.ferreira\Documents\jobs\07_anp\Test\in\"

Dim xlsFiles: xlsFiles = FilesInFolder(sourceFolder)

If UBound(xlsFiles) > 0 Then
	PivotTable2Workbook sourceFolder, resultFolder, xlsFiles
Else
	MsgBox "This folder does not exist or there is no xls/xlsx file in this folder!"
End If



Private Sub PivotTable2Workbook(sourceFolder, resultFolder, fileNames)

	Dim aHora : aHora = Now()
	Dim thisYear : thisYear = CStr(Year(aHora))
	Dim targetYear : targetYear = thisYear
    Dim foundYear : foundYear = False
	Dim thisMonth : thisMonth = Month(aHora)
    Const sheetName = "Plan1"
    Dim fileName
    
    For Each fileName In fileNames
    
    	Dim pgFldName, tableNames()
    	Dim file : file = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
    	If file = "9083" Then
    		pgFldName = "PRODUTO"
    		ReDim tableNames(1)
    		tableNames(0) = "Tabela dinâmica5"
    		tableNames(1) = "Tabela dinâmica7"
    	Else
    		ReDim tableNames(0)
    		If file = "1043" Then
    			pgFldName = "UN. DA FEDERAÇÃO"
    			tableNames(0) = "Tabela dinâmica1"
    		ElseIf file = "8476" Then
    			pgFldName = "ORIGEM"
    			tableNames(0) = "Tabela dinâmica5"
    		ElseIf file = "8740" Then
    			pgFldName = "PRODUTOR"
    			tableNames(0) = "Tabela dinâmica4"
    		ElseIf file = "11031" Then
    			pgFldName = "PRODUTO"
    			tableNames(0) = "Tabela dinâmica1"
    		End If
    	End If
    	
    	On Error Resume Next
    	MsgBox "opening: " & sourceFolder & fileName
    	Dim xlsWbk : Set xlsWbk = xlsApp.Workbooks.Open(sourceFolder & fileName, 0, True)
    	If Err.Number <> 0 Then ShowErr
    	
    	MsgBox fileName & " (" & pgFldName & ")" & vbCrLf & "tables:" & vbCrLf & Join(tableNames, vbCrLf)
    	
    	Dim tableName
    	For Each tableName In tableNames
    	
    		MsgBox tableName
    		Dim xlsApp : Set xlsApp = CreateObject("Excel.Application")
    		Set xlsApp.DisplayAlerts = False
			
			' Create result workbook and select first worksheet
            Dim xlsWbkNew : Set xlsWbkNew = xlsApp.Workbooks.Add
            Dim xlsWstNew : Set xlsWstNew = xlsWbkNew.Worksheets(1)
            xlsApp.Sheets(1).Select
            
        	Dim pvtTbl : pvtTbl = xlsWbk.Worksheets(sheetName).PivotTables(tableName)
        	MsgBox "IsNull: " & CStr(IsNull(pvtTbl))
        	Dim pgFld
        	
    		' Check for target subtitle
        	For Each pgFld In pvtTbl.PageFields
        		If pgFld.Name = pgFldName Then
					' Head of worksheet
                	xlsWstNew.Cells(1, 1).Value = "Mes"
                	xlsWstNew.Cells(1, 2).Value = pgFldName
                	xlsWstNew.Cells(1, 3).Value = "Valor"
                	xlsWstNew.Range("A1:C1").Font.Bold = True
                	' Set visible data from current year or past year
                	While foundYear = False
                		Dim pvtItm
                    	For Each pvtItm In pvtTbl.PivotFields("ANO").PivotItems
                        	If pvtItm.Name = targetYear Then
                            	pvtItm.Visible = True
                            	foundYear = True
                        	Else
                            	pvtItm.Visible = False
                        	End If
                    	Next
                    	If foundYear = False Then targetYear = thisYear - 1
                	Wend
                	' For every month
                	Dim i : i = 2
                	Dim j
                	For j = 1 To thisMonth
                    	' Filter by pattern in Pivot Item
                    	Dim pgPvtItm
                    	For Each pgPvtItm In pgFld.PivotItems
                        	pgFld.ClearAllFilters
                        	pgFld.CurrentPage = pgPvtItm.Name
                        	' Get pivot data from first month until current month for current year
                        	strMonth = MonthName(j)
                        	strMonthProp = UCase(Left(strMonth, 1)) & LCase(Right(strMonth, Len(strMonth) - 1))
                        	xlsWstNew.Cells(i, 1).Value = CheckMonth(j)
                        	xlsWstNew.Cells(i, 2).Value = pgPvtItm.Name
                        	xlsWstNew.Cells(i, 3).Value = pvtTbl.GetPivotData(CheckMonth(j), "Ano", targetYear).Value
                        	i = i + 1
                    	Next
                	Next
            	End If
            Next
            
            ' Default name for result workbook, e.g., 2016_ASFALTO.xlsx
            Dim xlsFilename : xlsFilename = targetYear & "_" & file & "_" & tableName & ".xlsx"
            ' Delete previous result workbook if exists
            If FileExists(resultFolder & xlsFilename) Then
            	FileDelete(resultFolder & xlsFilename)
            End If
            ' Save the result workbook
        	xlsWbkNew.SaveAs(folderPath & xlsFilename)
        	xlsWbkNew.Close
        	Set xlsWbkNew = Nothing
        	Set xlsWstNew = Nothing
        	' Finish the job
    		xlsWbk.Saved = True
    		xlsWbk.Close
    		xlsApp.Quit
    		Set xlsWbk = Nothing
    		Set xlsApp = Nothing
        Next
    Next
    
End Sub



' Return xls/xlsx files in existing folder
Private Function FilesInFolder(path)

	Dim arrFiles : arrFiles = Array()
	If FolderExists(path) Then
		Dim objFSO, objFolder, objFiles, objFile
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.GetFolder(path)
		Set objFiles = objFolder.Files
		For Each objFile In objFiles
			If objFSO.GetExtensionName(objFile) = "xls" Or objFSO.GetExtensionName(objFile) = "xlsx" Then
				Dim index : index = UBound(arrFiles)
				ReDim Preserve arrFiles(index + 1)
				arrFiles(index + 1) = objFile.Name
			End If
		Next
	End If
	FilesInFolder = arrFiles
	
End Function



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
Private Function FileExists(ByVal filePath)

   Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
   FileExists = fso.FileExists(filePath)
   Set fso = Nothing

End Function



' Receives a file path for check if exists and returns a boolean
Private Function FolderExists(ByVal folderPath)

   Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
   FolderExists = fso.FolderExists(folderPath)
   Set fso = Nothing

End Function



' Receives a file path for delete the file
Private Sub FileDelete(filePath)

    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile(filePath)
    Set fso = Nothing

End Sub