Sub SortByPlaces()
    Worksheets("Leltár").Sort.SortFields.Clear
    Range("A4:O235").Sort Key1:=Range("J4"), Key2:=Range("K4"), Order1:=xlAscending, _
    Order2:=xlDescending, Header:=xlYes
End Sub