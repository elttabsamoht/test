Sub Test()

    Dim Policko As New AutoArray
    
    Policko.Resize 1, 1
    Policko.Items(1, 1) = "A"
    
    MsgBox Policko.Items(1, 1), vbInformation

    Policko.Load (ActiveSheet.Range("A1:B2"))

    Policko = Nothing

End Sub

