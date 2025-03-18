Sub AccountantList()
    Dim lastRow As Long
    Dim cell As Range
    Dim newRow As Integer

    lastRow = Worksheets("Leltár").Range("A1048576").End(xlUp).Row
    newRow = 5
    
    For Each cell In Worksheets("Leltár").Range("E5:E" & lastRow).Cells
        Dim manuf As String
        Dim model As String
        Dim proc As String
        Dim memory As String
        Dim os As String
        Dim softwares As String
        Dim antivirus As String
        Dim location As String
        Dim db As String
        
        If cell.Value = "PC" Or cell.Value = "Laptop" Then
            manuf = Cells(cell.Row, "C").Value
            model = Cells(cell.Row, "D").Value
            proc = Cells(cell.Row, "G").Value
            memory = Cells(cell.Row, "H").Value
            os = Cells(cell.Row, "I").Value
            softwares = Cells(cell.Row, "J").Value
            antivirus = Cells(cell.Row, "K").Value
            location = Cells(cell.Row, "M").Value
            db = Cells(cell.Row, "N").Value
            
            Worksheets("Könyvvizsgálat").Cells(newRow, 1).Value = manuf
            Worksheets("Könyvvizsgálat").Cells(newRow, 2).Value = model
            Worksheets("Könyvvizsgálat").Cells(newRow, 3).Value = proc
            Worksheets("Könyvvizsgálat").Cells(newRow, 4).Value = memory
            Worksheets("Könyvvizsgálat").Cells(newRow, 5).Value = os
            Worksheets("Könyvvizsgálat").Cells(newRow, 6).Value = softwares
            Worksheets("Könyvvizsgálat").Cells(newRow, 7).Value = antivirus
            Worksheets("Könyvvizsgálat").Cells(newRow, 8).Value = location
            Worksheets("Könyvvizsgálat").Cells(newRow, 9).Value = db
            
            newRow = newRow + 1
        End If
    Next
    
    Debug.Print "A script lefutott"
    
End Sub
