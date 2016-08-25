' File: pivotTable2Workbook.vbs
' Description: Create workbook for pivot tables based on named columns (eg., month, product, value)
' Author: Artaxexes J. Ferreira
' Contact: artaxexes@ymail.com
' PivotTable ref.: https://msdn.microsoft.com/en-us/library/office/ff837611.aspx

Option Explicit



Const sourceFolder = "C:\Users\anddrei.ferreira\Documents\jobs\07_anp\Test\in\"
Const resultFolder = "C:\Users\anddrei.ferreira\Documents\jobs\07_anp\Test\out\"


Dim xlsFiles: xlsFiles = ListFilesInFolder(sourceFolder)

If UBound(xlsFiles) > 0 Then
	PivotTable2Workbook sourceFolder, resultFolder, xlsFiles
Else
	MsgBox "There is no xls/xlsx file in this folder"
End If



Private Sub PivotTable2Workbook(sourceFolder, resultFolder, fileNames)

	Dim aHora : aHora = Now()
	Dim thisYear : thisYear = CStr(Year(aHora))
	Dim targetYear : targetYear = thisYear
	Dim thisMonth : thisMonth = Month(aHora)
    Const sheetName = "Plan1"
    Dim fileName, tableNames()
    
    For Each fileName In fileNames

        Dim pgFldName
        Dim file : file = Mid(fileName, 1, InStrRev(fileName, ".") - 1)

    	If file = "9083" Then
    		pgFldName = "PRODUTO"
    		ReDim tableNames(1)
    		tableNames(0) = "Tabela dinâmica5"
    		tableNames(1) = "Tabela dinâmica7"
        ElseIf file = "8485" Then
            pgFldName = "PRODUTO"
            ReDim tableNames(4)
            tableNames(0) = "Tabela dinâmica1"
            tableNames(1) = "Tabela dinâmica2"
            tableNames(2) = "Tabela dinâmica3"
            tableNames(3) = "Tabela dinâmica7"
            tableNames(4) = "Tabela dinâmica12"
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
    	
        Dim xlsApp : Set xlsApp = CreateObject("Excel.Application")
		xlsApp.DisplayAlerts = False

    	On Error Resume Next
    	Dim xlsWbk : Set xlsWbk = xlsApp.Workbooks.Open(sourceFolder & fileName, 0, True)
    	If Err.Number <> 0 Then ShowErr
    	
    	Dim tableName
    	For Each tableName In tableNames
			
			' Create result workbook and select first worksheet
            Dim xlsWbkNew : Set xlsWbkNew = xlsApp.Workbooks.Add
            Dim xlsWstNew : Set xlsWstNew = xlsWbkNew.Worksheets(1)
            xlsApp.Sheets(1).Select
            
            Dim pvtTbl : Set pvtTbl = xlsWbk.Worksheets(sheetName).PivotTables(tableName)
        	Dim pgFld
            Dim foundYear : foundYear = False
        	
    		' Check for target subtitle
        	For Each pgFld In pvtTbl.PageFields
        		If pgFld.Name = pgFldName Then
					' Head of worksheet
                	xlsWstNew.Cells(1, 1).Value = "Mes"
                	xlsWstNew.Cells(1, 2).Value = pgFldName
                	xlsWstNew.Cells(1, 3).Value = "Valor"
                	xlsWstNew.Range("A1:C1").Font.Bold = True
                	' Set visible data from current year or past year
                    Dim pvtItm
                	While foundYear = False
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
                    Dim j, pgPvtItm
                	For j = 1 To thisMonth
                    	' Filter by pattern in Pivot Item
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
            Dim resultPath : resultPath = resultFolder & xlsFilename
            If FileExists(resultPath) Then
                FileDelete(resultPath)
            End If
            ' Save the result workbook
            xlsWbkNew.SaveAs(resultPath)
        	xlsWbkNew.Close
        	Set xlsWbkNew = Nothing
        	Set xlsWstNew = Nothing

        Next

        ' Finish the job
        xlsWbk.Saved = True
        xlsWbk.Close
        xlsApp.Quit
        Set xlsWbk = Nothing
        Set xlsApp = Nothing

    Next
    
End Sub



' Return xls/xlsx files in existing folder
Private Function ListFilesInFolder(path)

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
	Else
		MsgBox "This folder does not exist"
	End If
	ListFilesInFolder = arrFiles
	
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