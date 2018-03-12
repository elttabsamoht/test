'Sub test_basic_fns()
'
'    Dim testWB As Workbook
'    Dim testSavedWB As Workbook
'
'    Set testWB = openWB("D:\derpbox\actual\elttab\Dropbox\_code\vba\", "testik01.xlsx", False)
'    Set testSavedWB = openWB("D:\derpbox\actual\elttab\Dropbox\_code\vba\", "saved_testik01.xlsx")
'
'    Call saveWB(testWB, "D:\derpbox\actual\elttab\Dropbox\_code\vba\", "saved_testik01", "xlsx", True)
'
'    Call copyWS(testWB, testSavedWB, "copy", 0)
'
'    Dim zmenyStr As String
'
'    zmenyStr = "A>B;B>A"
'
'    Call moveCols(testWB, "copy", zmenyStr)
'
'End Sub


Sub optimizeVBA(isOn As Boolean)
    
    'Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    'ActiveSheet.DisplayPageBreaks = Not (isOn)
    
End Sub

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
            Case 1: 'rewriteOption by deleting sheet with same name in destination WB
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
    
    Dim userAnswer As Integer, wsNameNew As String
    wsNameNew = wsName & "_new"
           
    Do While (sheetExists(iWB, wsNameNew))
        
        userAnswer = MsgBox("Sheet name '" & wsNameNew & "' already exists. Overwrite?", vbQuestion & vbYesNo)
        
        If (userAnswer = vbYes) Then
            iWB.Sheets(wsNameNew).Delete
        Else
            Do
                wsNameNew = InputBox("Input new sheet name:", , wsNameNew)
            Loop While (wsNameNew = vbNullString)
        End If
            
    Loop
    
    Dim i As Long, colMovement() As String, colElements() As String, _
    SrcWS As Worksheet, DesWS As Worksheet, _
    SrcRng As Range, DesRng As Range
    
    colMovement = Split(colMovementStr, listDelimiter)
    
    Set SrcWS = iWB.Worksheets(wsName)
    iWB.Worksheets.Add After:=SrcWS
    Set DesWS = ActiveSheet: DesWS.Name = wsNameNew
    
    For i = 0 To UBound(colMovement)
        colElements() = Split(colMovement(i), colElementsDelimiter)
        Set SrcRng = SrcWS.Range(colElements(0) & ":" & colElements(0))
        Set DesRng = DesWS.Range(colElements(1) & ":" & colElements(1))
        'DesRng = SrcRng
        SrcRng.Copy DesRng
    Next i

End Sub
Public Sub moveColsWS(SrcWS As Worksheet, DesWS As Worksheet, colMovementStr As String, Optional listDelimiter As String = ";", Optional colElementsDelimiter As String = ">")
    
    Dim i As Long, colMovement() As String, colElements() As String, _
    SrcRng As Range, DesRng As Range
    
    colMovement = Split(colMovementStr, listDelimiter)
    
    For i = 0 To UBound(colMovement)
        colElements() = Split(colMovement(i), colElementsDelimiter)
        Set SrcRng = SrcWS.Range(colElements(0) & ":" & colElements(0))
        Set DesRng = DesWS.Range(colElements(1) & ":" & colElements(1))
        SrcRng.Copy DesRng
        
        'With DesRng
        '    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        '    .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        'End With

    Next i

End Sub

Public Function columnLetter(lngCol As Long) As String
    'columnLetter = Split(Cells(1, lngCol).Address, "$")(1) Split(Cells(1, lngCol).Address(True, False), "$")(0)
     columnLetter = Split(Cells(1, lngCol).Address(True, False, xlA1), "$")(0)

End Function

Public Function columnNumber(charCol As String) As Long
    columnNumber = Range(charCol & "1").Column
End Function

Public Function getLastCellRow(iWS As Worksheet, iRow As Long) As Long
    
    getLastCellRow = iWS.Cells(iRow, Columns.Count).End(xlToLeft).Column
    
End Function

Public Function getLastCellCol(iWS As Worksheet, iCol As String) As Long
    
    getLastCellCol = iWS.Cells(Rows.Count, columnNumber(iCol)).End(xlUp).Row
    
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

Public Sub filterOnArray(iWS As Worksheet, filterCol As String, filterKW() As String, Optional cancelOriginalFilter As Boolean = False, Optional cancelNewFilter As Boolean = False)
    'currently set up to only work with clean worksheet tables (header starting at A1); possible to do with optional range var and calculating the relative position of the column/field in the filter
    Dim fieldRange As String
    fieldRange = "A1:" & columnLetter(getLastCellRow(iWS, 1)) & getLastCellCol(iWS, "A")
    
    With iWS
    
        If (cancelOriginalFilter) Then
            .AutoFilterMode = False
            .Range(fieldRange).AutoFilter
        End If
    
        .Range(fieldRange).AutoFilter Field:=columnNumber(filterCol), Criteria1:=filterKW, Operator:=xlFilterValues

        If (cancelNewFilter) Then
            .AutoFilterMode = False
        End If

    End With

End Sub

'check other regular code and see if sub below can be wiped/replaced w filterOnArray (by using a 1 element string array)
Public Sub filter(iWS As Worksheet, filterCol As String, filterKW As String, Optional cancelOriginalFilter As Boolean = False, Optional cancelNewFilter As Boolean = False)
    'currently set up to only work with clean worksheet tables (header starting at A1); possible to do with optional range var and calculating the relative position of the column/field in the filter
    Dim fieldRange As String
    fieldRange = "A1:" & columnLetter(getLastCellRow(iWS, 1)) & getLastCellCol(iWS, "A")
    
    With iWS
    
        If (cancelOriginalFilter) Then
            .AutoFilterMode = False
            .Range(fieldRange).AutoFilter
        End If
    
        .Range(fieldRange).AutoFilter Field:=columnNumber(filterCol), Criteria1:=filterKW
        
        If (cancelNewFilter) Then
            .AutoFilterMode = False
        End If

    End With

End Sub

Public Sub filterOff(iWS As Worksheet)
    iWS.AutoFilterMode = False
End Sub

Public Sub sort(iWS As Worksheet, sortCol As String, Optional applySort As Boolean = True, Optional sortOrder As XlSortOrder = xlAscending, Optional cancelOriginalFilter As Boolean = False, Optional cancelNewFilter As Boolean = False)
    'currently set up to only work with clean worksheet tables (header starting at A1); possible to modify and if position of table to be sorted is specified
    'SortOn - sorting only by values, possible to modify input vars to include sort type (xlSortOnCellColor, xlSortOnFontColor)
    'DataOption = xlSortNormal - sorts numeric and text data separately; xlSortTextAsNumbers - treats text as numeric data for the sort.
    'could be transformed into a class and have proper methods (clearKey, applySort, addKey) / atm FilterOff as a workaround to not having multiple sorting key sets
    
    Dim fieldRange As String, sortRange As String, lLastRow As Long, strLastCol As String
    lLastRow = getLastCellCol(iWS, "A")
    strLastCol = columnLetter(getLastCellRow(iWS, 1))

    fieldRange = "A1:" & strLastCol & lLastRow
    sortRange = sortCol & "2:" & sortCol & lLastRow
    
    With iWS
    
        If (cancelOriginalFilter) Then
            .AutoFilterMode = False
            .Range(fieldRange).AutoFilter
            .AutoFilter.sort.SortFields.Clear
        Else:
            If (Not .AutoFilterMode) Then .Range(fieldRange).AutoFilter
        End If
        
        .AutoFilter.sort.SortFields.Add Key:=.Range(sortRange), SortOn:=xlSortOnValues, Order:=sortOrder, DataOption:=xlSortNormal
        
        If (applySort) Then
            With .AutoFilter.sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
        
        If (cancelNewFilter) Then
            .AutoFilterMode = False
        End If

    End With

End Sub


Public Sub copyPasteValuesWS(iWS As Worksheet, Optional clearColors As Boolean = True)

    Dim maxRow As Long, maxCol As Long
    maxRow = getLastRowWS(iWS)
    maxCol = getLastColWS(iWS)
    
    With iWS.Range(iWS.Cells(1, 1), iWS.Cells(maxRow, maxCol))
        .Copy
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Interior.ColorIndex = 0
    End With
    
End Sub


Public Sub deleteEmptyRowsWS(iWS As Worksheet, Optional keyColumn As Long = 1, Optional keyChar As String = "", Optional skipHeader As Boolean = True)
    
    Dim maxRow As Long, i As Long
    maxRow = getLastRowWS(iWS)

    i = CInt(Abs(skipHeader))
        
    Do
        If (iWS.Cells(i, keyColumn).Value = keyChar) Then
            iWS.Cells(i, keyColumn).EntireRow.Delete
            i = i - 1
            maxRow = maxRow - 1
        End If
        i = i + 1
    Loop While i <= maxRow
    

End Sub

