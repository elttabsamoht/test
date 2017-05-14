Attribute VB_Name = "TransposeDataCycle"
Sub TransposeDataCycle()

    Dim RowsForRecord As Integer: RowsForRecord = 0

    RowsForRecord = CInt(InputBox("Set number of rows for each item", "Rows per item", 0))
    If RowsForRecord = 0 Then Exit Sub

    Dim iRow As Long: iRow = 1
    Dim iCol As Long: iCol = 1
    Dim items As Long: items = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row / RowsForRecord
    Dim iCnt As Long

    For iCnt = 1 To items
        
        For iCol = 2 To RowsForRecord
            ActiveSheet.Cells(iRow, iCol).Value = ActiveSheet.Cells(iRow + iCol - 1, 1).Value
        Next iCol
    
        ActiveSheet.Range(Cells(iCnt + 1, 1), Cells(iRow + iCol - 2, 1)).Delete Shift:=xlUp
        iRow = iRow + 1
    
    Next iCnt

End Sub
