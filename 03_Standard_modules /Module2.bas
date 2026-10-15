Attribute VB_Name = "Module2"
' FUNKCJA POMOCNICZA DLA UKOŒNIKÓW
Private Function S() As String
    #If Mac Then
        S = "/"
    #Else
        S = "\"
    #End If
End Function

' PROCEDURA USUWANIA PACJENTA
Sub UsunPacjenta()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wsBaza As Worksheet
    Dim wierszDoUsuniecia As Long
    Dim ostatniWiersz As Long
    Dim pacjentFull As String
    Dim sciezkaPacjenta As String, sciezkaWydruki As String
    Dim nazwaFolderu As String
    
    On Error Resume Next
    Set wsBaza = Sheets("Baza_danych_medycznych")
    On Error GoTo 0
    
    wierszDoUsuniecia = ActiveCell.Row
    
    ' Sprawdzenie zakresu
    If Not Intersect(ActiveCell, ws.Range("D7:E1000")) Is Nothing Then
        If ws.Cells(wierszDoUsuniecia, "D").Value <> "" Then
            pacjentFull = ws.Cells(wierszDoUsuniecia, "D").Value & " " & ws.Cells(wierszDoUsuniecia, "E").Value
            nazwaFolderu = Replace(pacjentFull, " ", "_")
            
            ' Zapytanie o usuniêcie pacjenta z Excela
            If MsgBox("Czy na pewno chcesz usun¹æ pacjenta: " & pacjentFull & " z listy?", vbYesNo + vbExclamation) = vbYes Then
                
                ' Przygotowanie œcie¿ek folderów przy u¿yciu funkcji S()
                sciezkaPacjenta = ThisWorkbook.Path & S & "Dokumentacja_Pacjentow" & S & nazwaFolderu
                sciezkaWydruki = ThisWorkbook.Path & S & "Wydruki" & S & nazwaFolderu
                
                Application.EnableEvents = False
                Application.ScreenUpdating = False
                
                ' 1. Usuwanie z bazy medycznej
                If Not wsBaza Is Nothing Then
                    Dim fRng As Range
                    Set fRng = wsBaza.Columns("A").Find(pacjentFull, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not fRng Is Nothing Then fRng.EntireRow.Delete
                End If
                
                ' 2. Porz¹dkowanie listy g³ównej
                ostatniWiersz = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
                If wierszDoUsuniecia = ostatniWiersz Then
                    ws.Range("D" & wierszDoUsuniecia & ":E" & wierszDoUsuniecia).ClearContents
                Else
                    ws.Range("D" & wierszDoUsuniecia + 1 & ":E" & ostatniWiersz).Cut _
                        Destination:=ws.Range("D" & wierszDoUsuniecia)
                End If
                
                ' 3. Odœwie¿enie ramek
                Call RysujRamki
                
                Application.ScreenUpdating = True
                Application.EnableEvents = True
                
                ' SEKCJA 1: USUWANIE FOLDERU DOKUMENTACJI
                If FolderIstnieje(sciezkaPacjenta) Then
                    If MsgBox("Pacjent usuniêty z listy. Czy chcesz równie¿ USUN¥Æ folder z plikami medycznymi?", _
                              vbYesNo + vbQuestion, "Usuwanie dokumentacji") = vbYes Then
                        UsunPelnyFolder sciezkaPacjenta
                    End If
                End If
                
                ' SEKCJA 2: USUWANIE FOLDERU WYDRUKÓW
                If FolderIstnieje(sciezkaWydruki) Then
                    If MsgBox("Znaleziono folder z wygenerowanymi PDF (Wydruki). Czy chcesz go równie¿ USUN¥Æ?", _
                              vbYesNo + vbQuestion, "Usuwanie wydruków") = vbYes Then
                        UsunPelnyFolder sciezkaWydruki
                    End If
                End If
                
                MsgBox "Proces usuwania zakoñczony.", vbInformation
                
            End If
        Else
            MsgBox "Zaznaczona komórka jest pusta.", vbInformation
        End If
    Else
        MsgBox "Kliknij najpierw na Imiê lub Nazwisko w tabeli!", vbCritical
    End If
End Sub

' UNIWERSALNE FUNKCJE DO OBS£UGI PLIKÓW

Function FolderIstnieje(sciezka As String) As Boolean
    FolderIstnieje = (Dir(sciezka, vbDirectory) <> "")
End Function

Sub UsunPelnyFolder(sciezka As String)
    On Error Resume Next
    #If Mac Then
        ' Na Macu u¿ywam AppleScript, aby przenieœæ folder do kosza
        Dim script As String
        script = "tell application ""Finder"" to delete POSIX file """ & sciezka & """"
        MacScript script
    #Else
        ' Na Windowsie u¿ywam tradycyjnego FileSystemObject
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFolder sciezka, True
    #End If
    On Error GoTo 0
End Sub

