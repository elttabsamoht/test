Sub Sort(iWS As Worksheet, sortCol As String, Optional applySort As Boolean = True, Optional sortOrder As XlSortOrder = xlAscending, Optional cancelOriginalFilter As Boolean = False, Optional cancelNewFilter As Boolean = False)
    'currently set up to only work with clean worksheet tables (header starting at A1); possible to modify and if position of table to be sorted is specified
    'SortOn - sorting only by values, possible to modify input vars to include sort type (xlSortOnCellColor, xlSortOnFontColor)
    'DataOption = xlSortNormal - sorts numeric and text data separately; xlSortTextAsNumbers - treats text as numeric data for the sort.
    'could be transformed into a class and have proper methods (clearKey, applySort, addKey)
    
    Dim fieldRange As String, sortRange As String, lLastRow As Long, strLastCol As String
    lLastRow = getLastCellCol(iWS, "A")
    strLastCol = columnLetter(getLastCellRow(iWS, 1))

    fieldRange = "A1:" & strLastCol & lLastRow
    sortRange = sortCol & "2:" & sortCol & lLastRow
    
    With iWS
    
        If (cancelOriginalFilter) Then
            .AutoFilterMode = False
            .Range(fieldRange).AutoFilter
            .AutoFilter.Sort.SortFields.Clear
        Else:
            If (Not .AutoFilterMode) Then .Range(fieldRange).AutoFilter
        End If
        
        .AutoFilter.Sort.SortFields.Add Key:=.Range(sortRange), SortOn:=xlSortOnValues, Order:=sortOrder, DataOption:=xlSortNormal
        
        If (applySort) Then
            With .AutoFilter.Sort
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