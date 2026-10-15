Attribute VB_Name = "Module1"
' FUNKCJA POMOCNICZA DLA UKOŚNIKÓW
Private Function S() As String
    #If Mac Then
        S = "/"
    #Else
        S = "\"
    #End If
End Function

' PROCEDURA DODAWANIA PACJENTA
    Dim ws As Worksheet: Set ws = Sheets("Pacjenci")
    Dim imie As String, nazwisko As String
    Dim nastepnyWiersz As Long
    Dim sciezkaGlowna As String, sciezkaPacjenta As String
    
    imie = Trim(ws.Range("J6").Value)
    nazwisko = Trim(ws.Range("J7").Value)
    
    ' Walidacja danych wejściowych
    If imie = "" Or nazwisko = "" Then
        MsgBox "Wpisz imię w J6 i nazwisko w J7!", vbExclamation
        Exit Sub
    End If
    
    ' LOGIKA FOLDERÓW
    sciezkaGlowna = ThisWorkbook.Path & S & "Dokumentacja_Pacjentow"
    sciezkaPacjenta = sciezkaGlowna & S & imie & "_" & nazwisko
    
    On Error Resume Next
    ' Sprawdzenie i tworzenie folderu głównego
    If Dir(sciezkaGlowna, vbDirectory) = "" Then
        MkDir sciezkaGlowna
    End If
    
    ' Sprawdzenie i tworzenie folderu pacjenta
    If Dir(sciezkaPacjenta, vbDirectory) = "" Then
        MkDir sciezkaPacjenta
    End If
    On Error GoTo 0

    ' Wyłączamy zdarzenia (blokada Arkusza przed komunikatem Undo)
    Application.EnableEvents = False
    
    ' Szukanie wiersza
    nastepnyWiersz = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row + 1
    If nastepnyWiersz < 7 Then nastepnyWiersz = 7
    
    ' Wpisywanie danych
    ws.Cells(nastepnyWiersz, "D").Value = imie
    ws.Cells(nastepnyWiersz, "E").Value = nazwisko
    
    ' Czyszczenie pól i rysowanie ramek
    ws.Range("J6:J7").ClearContents
    
    ' Upewnienie się, że procedura RysujRamki też nie ma błędów
    Call RysujRamki
    
    ' Włączamy zdarzenia
    Application.EnableEvents = True
    
    MsgBox "Pacjent dodany i folder utworzony!", vbInformation
End Sub
