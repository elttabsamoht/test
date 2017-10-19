Attribute VB_Name = "lib_basic_fns_v2"
'Attribute VB_Name = "mod_functions_beta"
'Sub test_basic_fns()
'
'    Dim testWB As Workbook
'    'Dim testSavedWB As Workbook
'
'    Set testWB = openWB("D:\derpbox\actual\elttab\Dropbox\_code\vba\", "testik01.xlsx", False)
'    'Set testSavedWB = openWB("D:\derpbox\actual\elttab\Dropbox\_code\vba\", "saved_testik01.xlsx")
'
'    'Call saveWB(testWB, "D:\derpbox\actual\elttab\Dropbox\_code\vba\", "saved_testik01", "xlsx", True)
'
'    'Call copyWS(testWB, testSavedWB, "copy", 0)
'
'    Dim zmenyStr As String
'
'    zmenyStr = "A>B;B>A"
'
'    Call moveCols(testWB, "copy", zmenyStr)
'
'End Sub

Public Sub saveWB(iWB As Workbook, FilePath As String, FileName As String, fileFormat As String, Optional closeAfterSaving As Boolean = True)
    
    Dim fileFormatNum As Integer
    
    Application.DisplayAlerts = False
    
    Select Case fileFormat
        Case "xlsx": fileFormatNum = 51
        Case "xlsm": fileFormatNum = 52
        Case "xls": fileFormatNum = 56
        Case "csv": fileFormatNum = 6
        Case "txt": fileFormatNum = -4158
        
    End Select
    
    On Error GoTo errHandler:
    
    iWB.SaveAs FilePath & FileName & "." & fileFormat, fileFormatNum
    
    If (closeAfterSaving) Then iWB.Close
    
    Application.DisplayAlerts = True
    
    Exit Sub
    
errHandler:
Application.DisplayAlerts = True
MsgBox "Unable to save file", vbCritical
    
End Sub

Public Function openWB(FilePath As String, FileName As String, Optional readOnly As Boolean = True, Optional calculationsOff As Boolean = False, Optional updateLinksOff As Boolean = True) As Workbook
    
'disable macros ako specialna moznost? - disable events iba?

    If (calculationsOff) Then Application.Calculation = xlManual '!kalkulacie ostavaju vypnute - staci to tu?

    Dim updateLinksNum As Integer

    If (updateLinksOff) Then
        updateLinksNum = 0
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    Else
        updateLinksNum = 1
    End If
    
    Set openWB = Workbooks.Open(FilePath & FileName, updateLinksNum, readOnly)

    If (updateLinksOff) Then
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If

End Function

Public Function sheetExists(iWB As Workbook, wsName As String) As Boolean
      
    sheetExists = False
    Dim i As Integer
      
    For i = 1 To iWB.Worksheets.Count
        If iWB.Worksheets(i).Name = wsName Then
            sheetExists = True
        End If
    Next i
      
End Function


Public Sub copyWS(srcWB As Workbook, DesWB As Workbook, wsName As String, Optional rewriteOption As Integer = 1, Optional moveOnly As Boolean = False)

    'currently unable to copy multiple WS
    'currently unable to copy into a new WS

    Application.DisplayAlerts = False
    
    Dim wsNameSuffix As String: wsNameSuffix = ""
     
    If (sheetExists(DesWB, wsName)) Then
        Select Case rewriteOption
            Case 1: 'rewriteOption by deleting shet with same name in destination WB
                DesWB.Sheets(wsName).Delete
            Case 0: 'add a "_new" suffix to the copied worksheet
                wsNameSuffix = "_new"
            Case 2: 'rename the existing sheet in the destination workbook by adding the "_old" suffix
                DesWB.Sheets(wsName).Name = wsName & "_old"
        End Select
    End If
    
    srcWB.Sheets(wsName).Copy After:=DesWB.Sheets(DesWB.Sheets.Count)
    ActiveSheet.Name = wsName & wsNameSuffix
    
    If (moveOnly) Then
        srcWB.Sheets(wsName).Delete
    End If
    
    Application.DisplayAlerts = True
    
End Sub

Public Sub moveCols(iWB As Workbook, wsName As String, colMovementStr As String, Optional listDelimiter As String = ";", Optional colElementsDelimiter As String = ">")
    
    Dim i As Long, colMovement() As String, colElements() As String, _
    srcWS As Worksheet, DesWS As Worksheet, _
    SrcRng As Range, DesRng As Range
    
    colMovement = Split(colMovementStr, listDelimiter)
    
    Set srcWS = iWB.Worksheets(wsName): iWB.Worksheets.Add After:=srcWS 'After:=iWB.Worksheets(wsName)
    Set DesWS = ActiveSheet: DesWS.Name = wsName & "_new"
    
    For i = 0 To UBound(colMovement)
        colElements() = Split(colMovement(i), colElementsDelimiter)
        Set SrcRng = srcWS.Range(colElements(0) & ":" & colElements(0))
        Set DesRng = DesWS.Range(colElements(1) & ":" & colElements(1))
        'desRng = srcRng
        SrcRng.Copy DesRng
    Next i

End Sub

Public Function colLetter(lngCol As Long) As String
    colLetter = Split(Cells(1, lngCol).Address, "$")(1) 'Split(Cells(1, lngCol).Address(True, False), "$")(0)
End Function

Public Function colNumber(charCol As String) As Long
    colNumber = Range(ColName & 1).Column
End Function


Public Function getLastRowWS(iWS As Worksheet) As Long

    With iWS
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            getLastRowWS = .Cells.Find(What:="*", _
                          After:=.Range("A1"), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          searchdirection:=xlPrevious, _
                          MatchCase:=False).Row
        Else
            getLastRowWS = 1
        End If
    End With

End Function

Public Function getLastColWS(iWS As Worksheet) As Long

    With iWS
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            getLastColWS = .Cells.Find(What:="*", _
                          After:=.Range("A1"), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, _
                          searchdirection:=xlPrevious, _
                          MatchCase:=False).Column
        Else
            getLastColWS = 1
        End If
    End With

End Function




Public Sub loadTextData(FileName As String, iWS As Worksheet)
   
    FileName = "Z:\My Documents\!work\code\vba\FPR\test1_data_full.txt"
        
'    Workbooks.OpenText Filename:=FilePath, _
'    DataType:=xlDelimited, Tab:=True
'
'    Set LoadWB = ThisWorkbook
    
    
    Workbooks.OpenText FileName:=FileName, Origin:=xlWindows, _
    StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=True, Comma:=False, _
    Space:=False, Other:=False
                 
    Set LoadWB = ActiveWorkbook
                 
    Set LoadWS = ActiveSheet

'    Workbooks.Open ("C:\Test.xls"), False
'
'    Workbooks(MyBook).Sheets(MySheet).Cells.Copy Workbooks("Test.xls").Sheets("Sheet1").Range("A1")
'    Workbooks(MyBook).Close False

End Sub

Sub FilterAndSort(FilterCol As String)
    
    'FilterCol = "E"
    
    With ActiveSheet
        
        .AutoFilterMode = False
        .Range("A1", Range("XFD1").End(xlToLeft)).Select
        .Range(Selection, Range("A" & Rows.Count).End(xlUp)).AutoFilter
        .AutoFilter.Sort.SortFields.Add Key:=Range(FilterCol & "1:" & FilterCol & getLastRowWS(ActiveSheet)), SortOn:=xlSortOnValues, Order:=xlAscending
        .AutoFilter.Sort.Header = xlYes
        .AutoFilter.Sort.Apply
        .AutoFilterMode = False
    End With
    
End Sub

'Sub Filter(FilterCol As String, FilterKW As String, Optional LeaveFilter As Boolean = False)
'
'    With ActiveSheet
'
'        .AutoFilterMode = False
'        .Range("A1", Range("XFD1").End(xlToLeft)).Select
'
'        With .Range(Selection, Range("A" & Rows.Count).End(xlUp))
'            .AutoFilter
'            .AutoFilter Field:=ColumnNumber(FilterCol), Criteria1:=FilterKW
'        End With
'
'        '.AutoFilter.Sort.SortFields.Add Key:=Range(FilterCol & "1:" & FilterCol & GetLastRowWS(ActiveSheet)), SortOn:=xlSortOnValues, Order:=xlAscending
'        '.AutoFilter.Sort.Header = xlYes
'        '.AutoFilter.Sort.Apply
'        If (LeaveFilter = False) Then
'        .AutoFilterMode = False
'        End If
'
'    End With
'
'
'End Sub


Sub CutAndSave(iWS As Worksheet, FilterCol As String, FilePath As String, FileName As String)

'zatial sprav cez loop a cut default na to, co je v bunke (= kod fondu), prerob neskor cez filtrovanie na vytiahnutie len dat, ktore su vo filtri

'   Do
 
    
'   Loop Until (iWS.Range("A2").Value = "")
   
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
    
    'DesWB.SaveAs FileName:="Z:\My Documents\!work\code\vba\FPR\test.xls", FileFormat:=56
    DesWB.SaveAs FileName:=FilePath & FileName ', FileFormat:=56
    
    
    Application.DisplayAlerts = True
    
    
    Set DesWB = Nothing
    Set DesWS = Nothing

    
End Sub
