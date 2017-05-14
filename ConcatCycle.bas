Attribute VB_Name = "ConcatCycle"
Function ConcatCycle(tCycles As Integer, tRng As Range, Optional tSpace As String = " ") As String

    Dim iString As String: iString = ""
    Dim iCnt As Integer

    For iCnt = 1 To tCycles
        ConcatCycle = ConcatCycle & tRng.Cells(1, iCnt).Value
        If (iCnt <> tCycles) Then ConcatCycle = ConcatCycle & " "
    Next iCnt

End Function
