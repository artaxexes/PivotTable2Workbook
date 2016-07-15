' File: pivotTable2Workbook.vbs
' Description: Create workbook for pivot tables based on named columns (eg., month, product, value)
' Author: Artaxexes J. Ferreira
' Contact: artaxexes@ymail.com
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
    
    Dim xlsWbkNew
    Dim xlsWstNew
    Dim xlsFilename
    Dim PvtTbl
    Dim PgFld
    Dim PvtItm
    Dim PgPvtItm
    Dim i
    Dim j
    
    Dim thisYear : thisYear = CStr(Year(Now()))
    Dim thisMonth : thisMonth = Month(Now())

    Dim foundYear : foundYear = False
    Dim targetYear : targetYear = thisYear

    Dim xlsApp : Set xlsApp = CreateObject("Excel.Application")
    xlsApp.DisplayAlerts = False

    On Error Resume Next
    Dim xlsWbk : Set xlsWbk = xlsApp.Workbooks.Open(xlsPath, 0, True) 'xlSrcWbk
    If Err.Number <> 0 Then ShowErr

    ' For each pivot table: check and filter by pattern, set visible for current year, create a workbook with filtered data
    For Each PvtTbl In xlsWbk.Worksheets("Plan1").PivotTables
        ' Check for pattern in Page Item
        For Each PgFld In PvtTbl.PageFields
            If PgFld.Name = "PRODUTO" Then
                ' Create result workbook and select first worksheet
                Set xlsWbkNew = xlsApp.Workbooks.Add
                Set xlsWstNew = xlsWbkNew.Worksheets(1)
                xlsApp.Sheets(1).Select
                ' Head of current worksheet
                xlsWstNew.Cells(1, 1).Value = "Mes"
                xlsWstNew.Cells(1, 2).Value = "Produto"
                xlsWstNew.Cells(1, 3).Value = "Valor"
                xlsWstNew.Range("A1:C1").Font.Bold = True
                ' Set visible data from current year or past year
                While foundYear = False
                    For Each PvtItm In PvtTbl.PivotFields("ANO").PivotItems
                        If PvtItm.Name = targetYear Then
                            PvtItm.Visible = True
                            foundYear = True
                        Else
                            PvtItm.Visible = False
                        End If
                    Next
                    If foundYear = False Then targetYear = thisYear - 1
                WEnd
                ' For every month
                i = 2
                For j = 1 To thisMonth
                    ' Filter by pattern in Pivot Item
                    For Each PgPvtItm In PgFld.PivotItems
                        PgFld.ClearAllFilters
                        PgFld.CurrentPage = PgPvtItm.Name
                        ' Get pivot data from first month until current month for current year
                        'strMonth = MonthName(j)
                        'strMonthProp = UCase(Left(strMonth, 1)) & LCase(Right(strMonth, Len(strMonth) - 1))
                        xlsWstNew.Cells(i, 1).Value = CheckMonth(j)
                        xlsWstNew.Cells(i, 2).Value = PgPvtItm.Name
                        xlsWstNew.Cells(i, 3).Value = PvtTbl.GetPivotData(CheckMonth(j), "Ano", targetYear).Value
                        i = i + 1
                    Next
                Next
                ' Default name for result workbook, e.g., 2016_ASFALTO.xlsx
                xlsFilename = targetYear & "_" & PvtTbl.Name & ".xlsx"
                ' Delete previous result workbook if exists
                If FileExists(folderPath & xlsFilename) Then
                    FileDelete(folderPath & xlsFilename)
                End If
            End If
        Next
        ' Save the result workbook
        xlsWbkNew.SaveAs(folderPath & xlsFilename)
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