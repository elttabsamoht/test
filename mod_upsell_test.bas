Attribute VB_Name = "mod_upsell_test"
Sub upsell_test()

'podmienka;ID bla;key ID;nieco ine;cas
';8;a;bla;10.9.2017 10:30:30
';1;a;bla;12.9.2017 10:30:40
';5;b;bla;10.9.2017 10:30:40
';2;a;bla;10.9.2017 10:30:40
';3;a;bla;10.5.2017 10:30:40
';4;b;bla;10.9.2017 10:30:30


    Dim upsellIBMA As New AutoArray, upsellWS As Worksheet, lastRow As Long
    Set upsellWS = ActiveSheet
    
    lastRow = Range("D" & Rows.Count).End(xlUp).Row
    
    With Range("A2", Range("E" & Rows.Count).End(xlUp))
    .Sort Key1:=.Cells(1, 3), Order1:=xlAscending, _
           key2:=.Cells(1, 5), order2:=xlAscending ', _
           'key3:=.Cells(1, 1), order3:=xlAscending, Header:=xlGuess
    End With
    
    
    'upsellIBMA.Resize 2, 7
    upsellIBMA.Load upsellWS.Range("C2:C7, D2:D7") ' & lastRow & ",D2:D" & lastRow)

    MsgBox upsellIBMA.Items(1, 1)

    Set upsellIBMA = Nothing

End Sub


