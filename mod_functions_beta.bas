Attribute VB_Name = "mod_functions_beta"
'Sub test()
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



Sub saveWB(iWB As Workbook, filePath As String, fileName As String, fileFormat As String, Optional closeAfterSaving As Boolean = True)
    
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
    
    iWB.SaveAs filePath & fileName & "." & fileFormat, fileFormatNum
    
    If (closeAfterSaving) Then iWB.Close
    
    Application.DisplayAlerts = True
    
    Exit Sub
    
errHandler:
Application.DisplayAlerts = True
MsgBox "Unable to save file", vbCritical
    
End Sub

Function openWB(filePath As String, fileName As String, Optional readOnly As Boolean = True, Optional calculationsOff As Boolean = False, Optional updateLinksOff As Boolean = True) As Workbook
    
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
    
    Set openWB = Workbooks.Open(filePath & fileName, updateLinksNum, readOnly)

    If (updateLinksOff) Then
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If

End Function

Function sheetExists(iWB As Workbook, wsName As String) As Boolean
      
    sheetExists = False
    Dim i As Integer
      
    For i = 1 To iWB.Worksheets.Count
        If iWB.Worksheets(i).Name = wsName Then
            sheetExists = True
        End If
    Next i
      
End Function


Sub copyWS(srcWB As Workbook, desWB As Workbook, wsName As String, Optional rewriteOption As Integer = 1, Optional moveOnly As Boolean = False)

    'currently unable to copy multiple WS
    'currently unable to copy into a new WS

    Application.DisplayAlerts = False
    
    Dim wsNameSuffix As String: wsNameSuffix = ""
     
    If (sheetExists(desWB, wsName)) Then
        Select Case rewriteOption
            Case 1: 'rewriteOption by deleting shet with same name in destination WB
                desWB.Sheets(wsName).Delete
            Case 0: 'add a "_new" suffix to the copied worksheet
                wsNameSuffix = "_new"
            Case 2: 'rename the existing sheet in the destination workbook by adding the "_old" suffix
                desWB.Sheets(wsName).Name = wsName & "_old"
        End Select
    End If
    
    srcWB.Sheets(wsName).Copy After:=desWB.Sheets(desWB.Sheets.Count)
    ActiveSheet.Name = wsName & wsNameSuffix
    
    If (moveOnly) Then
        srcWB.Sheets(wsName).Delete
    End If
    
    Application.DisplayAlerts = True
    
End Sub

Sub moveCols(iWB As Workbook, wsName As String, colMovementStr As String, Optional listDelimiter As String = ";", Optional colElementsDelimiter As String = ">")
    
    Dim i As Long, colMovement() As String, colElements() As String, _
    srcWS As Worksheet, desWS As Worksheet, _
    srcRng As Range, desRng As Range
    
    colMovement = Split(colMovementStr, listDelimiter)
    
    Set srcWS = iWB.Worksheets(wsName): iWB.Worksheets.Add After:=srcWS 'After:=iWB.Worksheets(wsName)
    Set desWS = ActiveSheet: desWS.Name = wsName & "_new"
    
    For i = 0 To UBound(colMovement)
        colElements() = Split(colMovement(i), colElementsDelimiter)
        Set srcRng = srcWS.Range(colElements(0) & ":" & colElements(0))
        Set desRng = desWS.Range(colElements(1) & ":" & colElements(1))
        'desRng = srcRng
        srcRng.Copy desRng
    Next i

End Sub

Function colLetter(lngCol As Long) As String
    colLetter = Split(Cells(1, lngCol).Address, "$")(1) 'Split(Cells(1, lngCol).Address(True, False), "$")(0)
End Function

Function colNumber(charCol As String) As Long
    colNumber = Range(ColName & 1).Column
End Function
