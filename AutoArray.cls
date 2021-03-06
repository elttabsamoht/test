'class AutoArray
'Attributes

Private pItems() As String
Private pRowMax As Integer
Private pColMax As Integer

Public Property Get Items(iRow As Integer, iCol As Integer) As String
    Items = pItems(iRow, iCol)
End Property

Public Property Let Items(iRow As Integer, iCol As Integer, iStr As String)
    pItems(iRow, iCol) = iStr
End Property

Public Property Get RowMax() As Integer
    RowMax = pRowMax
End Property

Public Property Let RowMax(iRow As Integer)
    pRowMax = iRow
End Property
Public Property Get ColMax() As Integer
    ColMax = pColMax
End Property

Public Property Let ColMax(iCol As Integer)
    pColMax = iCol
End Property

Public Sub Resize(iRow As Integer, iCol As Integer)
    Let RowMax = iRow
    Let ColMax = iCol
    ReDim pItems(pRowMax,pColMax)
End Sub

Public Sub Load(iRng As Range)
    Let RowMax = iRng.Rows.Count: Let ColMax = iRng.Columns.Count
    Resize pRowMax, pColMax
    
    Dim iRow As Integer: Dim iCol As Integer
    Dim iStr As String: iStr = ""
    
    For iRow = 1 To pRowMax
        For iCol = 1 To pColMax
            pItems(iRow, iCol) = iRng.Cells(iRow, iCol).Value
            iStr = iStr & pItems(iRow, iCol) & " "
        Next iCol
        iStr = iStr & vbCrLf
    Next iRow
    
    MsgBox iStr, vbInformation

End Sub

Private Sub Class_Initialize()
    ReDim pItems(0, 0)
    Let RowMax = 0: Let ColMax = 0
    Resize pRowMax, pColMax
End Sub

Private Sub Class_Terminate()

End Sub
