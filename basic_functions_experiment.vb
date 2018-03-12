Attribute VB_Name = "mod_functions_beta"

Sub test_basic_fns()
	
    Dim testWB As Workbook
    Dim testSavedWB As Workbook

    Set testWB = openWB("D:\derpbox\actual\elttab\Dropbox\_code\vba\", "testik01.xlsx", False)
    Set testSavedWB = openWB("D:\derpbox\actual\elttab\Dropbox\_code\vba\", "saved_testik01.xlsx")

    Call saveWB(testWB, "D:\derpbox\actual\elttab\Dropbox\_code\vba\", "saved_testik01", "xlsx", True)

    Call copyWS(testWB, testSavedWB, "copy", 0)

    Dim zmenyStr As String

    zmenyStr = "A>B;B>A"

    Call moveCols(testWB, "copy", zmenyStr)

End Sub

Public Sub loadTextData(FileName As String, iWS As Worksheet)
   
    FileName = "Z:\My Documents\!work\code\vba\FPR\test1_data_full.txt"

    Workbooks.OpenText Filename:=FilePath, _
    DataType:=xlDelimited, Tab:=True

    Set LoadWB = ThisWorkbook
    
    
    Workbooks.OpenText FileName:=FileName, Origin:=xlWindows, _
    StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=True, Comma:=False, _
    Space:=False, Other:=False
 
    Set LoadWB = ActiveWorkbook

    Set LoadWS = ActiveSheet

    Workbooks.Open ("C:\Test.xls"), False

    Workbooks(MyBook).Sheets(MySheet).Cells.Copy Workbooks("Test.xls").Sheets("Sheet1").Range("A1")
    Workbooks(MyBook).Close False

End Sub

Sub FilterAndSort(iWS as worksheet, FilterCol As String)
    
    'FilterCol = "E"
    
    With iWS
        
        .AutoFilterMode = False
        .Range("A1", Range("XFD1").End(xlToLeft)).Select
        .Range(Selection, Range("A" & Rows.Count).End(xlUp)).AutoFilter
        .AutoFilter.Sort.SortFields.Add Key:=Range(FilterCol & "1:" & FilterCol & getLastRowWS(ActiveSheet)), SortOn:=xlSortOnValues, Order:=xlAscending
        .AutoFilter.Sort.Header = xlYes
        .AutoFilter.Sort.Apply
        .AutoFilterMode = False
    End With

End Sub

Sub CutAndSave(iWS As Worksheet, FilterCol As String, FilePath As String, FileName As String)
	
	'zatial sprav cez loop a cut default na to, co je v bunke (= kod fondu), prerob neskor cez filtrovanie na vytiahnutie len dat, ktore su vo filtri
	
   Do
 
    
   Loop Until (iWS.Range("A2").Value = "")
   
    Dim iFilterKW_LastRow As Long
    Dim iLastCol As Long
    Dim DesWB As Workbook
    Dim DesWS As Worksheet
    Dim DesRng As Range: Dim SrcRng As Range
    
    iFilterKW_LastRow = Range(iWS.Cells(2, FilterCol), iWS.Cells(GetLastCellCol(iWS, FilterCol), FilterCol)).Find(iWS.Cells(2, FilterCol).Value, searchdirection:=xlPrevious).Row
    iLastCol = GetLastCellRow(iWS, 1)
    
    MsgBox iFilterKW_LastRow
    
    Set DesWB = Workbooks.Add

    
    Application.DisplayAlerts = False
    
    DesWB.Sheets(2).Delete: DesWB.Sheets(2).Delete:
    Set DesWS = DesWB.Sheets("Sheet1"): DesWS.Name = "data"
    
    Set DesRng = DesWS.Range(DesWS.Cells(1, 1), DesWS.Cells(iFilterKW_LastRow - 1, iLastCol))
    Set SrcRng = iWS.Range(iWS.Cells(2, 1), iWS.Cells(iFilterKW_LastRow, iLastCol))
    
    DesRng.Value = SrcRng.Value
    
    DesWB.SaveAs FileName:="Z:\My Documents\!work\code\vba\FPR\test.xls", FileFormat:=56
    DesWB.SaveAs FileName:=FilePath & FileName , FileFormat:=56
    
    
    Application.DisplayAlerts = True
    
    
    Set DesWB = Nothing
    Set DesWS = Nothing

    
End Sub

Sub vycisti()
    
    Application.ScreenUpdating = False

    Dim numOfRows As Long, iRow As Long
    numOfRows = getLastRowWS(ActiveSheet)
    
    With ActiveSheet
        For iRow = 2 To numOfRows
    
            If (.Cells(iRow, 7).Value = 0) Then 'chabe podmienky
                If (.Cells(iRow, 8).Value = 0) Then
                    .Cells(iRow, 1).EntireRow.Delete
                    numOfRows = numOfRows - 1
                    iRow = iRow - 1
                End If
            End If
        Next iRow
    End With
    
    

    Application.ScreenUpdating = True

End Sub

Option Explicit
 
Sub ExcelDiet()
     
    Dim j               As Long
    Dim k               As Long
    Dim LastRow         As Long
    Dim LastCol         As Long
    Dim ColFormula      As Range
    Dim RowFormula      As Range
    Dim ColValue        As Range
    Dim RowValue        As Range
    Dim Shp             As Shape
    Dim ws              As Worksheet
     
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     
    On Error Resume Next
     
    For Each ws In Worksheets
        With ws
             'Find the last used cell with a formula and value
             'Search by Columns and Rows
            On Error Resume Next
            Set ColFormula = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            Set ColValue = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            Set RowFormula = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            Set RowValue = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            On Error GoTo 0
             
             'Determine the last column
            If ColFormula Is Nothing Then
                LastCol = 0
            Else
                LastCol = ColFormula.Column
            End If
            If Not ColValue Is Nothing Then
                LastCol = Application.WorksheetFunction.Max(LastCol, ColValue.Column)
            End If
             
             'Determine the last row
            If RowFormula Is Nothing Then
                LastRow = 0
            Else
                LastRow = RowFormula.Row
            End If
            If Not RowValue Is Nothing Then
                LastRow = Application.WorksheetFunction.Max(LastRow, RowValue.Row)
            End If
             
             'Determine if any shapes are beyond the last row and last column
            For Each Shp In .Shapes
                j = 0
                k = 0
                On Error Resume Next
                j = Shp.TopLeftCell.Row
                k = Shp.TopLeftCell.Column
                On Error GoTo 0
                If j > 0 And k > 0 Then
                    Do Until .Cells(j, k).Top > Shp.Top + Shp.Height
                        j = j + 1
                    Loop
                    If j > LastRow Then
                        LastRow = j
                    End If
                    Do Until .Cells(j, k).Left > Shp.Left + Shp.Width
                        k = k + 1
                    Loop
                    If k > LastCol Then
                        LastCol = k
                    End If
                End If
            Next
             
            .Range(.Cells(1, LastCol + 1), .Cells(.Rows.Count, .Columns.Count)).EntireColumn.Delete
            .Range("A" & LastRow + 1 & ":A" & .Rows.Count).EntireRow.Delete
        End With
    Next
     
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Done", vbInformation
     
End Sub


