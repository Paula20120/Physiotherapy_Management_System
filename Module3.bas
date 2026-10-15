Attribute VB_Name = "Module3"
Sub RysujRamki()
    Dim ws As Worksheet: Set ws = Sheets("Pacjenci")
    Dim ostatniWiersz As Long
    
    ' Czyszczenie starych ramek
    ws.Range("D7:E1000").Borders.LineStyle = xlNone
    
    ' Szukamy ostatniego pacjenta
    ostatniWiersz = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ' Rysujemy nowe ramki tylko dla wype³nionych wierszy
    If ostatniWiersz >= 7 Then
        With ws.Range("D7:E" & ostatniWiersz).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
End Sub
