Attribute VB_Name = "CountChar"
Function CountChar(tString As String, tChar As String) As Long
    CountChar = Len(tString) - Len(Replace(tString, tChar, ""))
End Function
