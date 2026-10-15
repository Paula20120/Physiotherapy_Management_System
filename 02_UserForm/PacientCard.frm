VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKartaPacjenta 
   Caption         =   "Karta Pacjenta"
   ClientHeight    =   10488
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14268
   OleObjectBlob   =   "PacientCard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKartaPacjenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rIndex As Long

' --- FUNKCJA UNIWERSALNEGO UKOŒNIKA ---
Public Function S() As String
    #If Mac Then
        S = "/"
    #Else
        S = "\"
    #End If
End Function

Private Sub txtWywiad_Change()

End Sub

' 1. INICJALIZACJA FORMULARZA
Private Sub UserForm_Initialize()
    With cmbStatusPlatnosci
        .Clear
        .AddItem "Op³acono"
        .AddItem "Nieop³acono"
        .Value = "Nieop³acono"
    End With
End Sub

' 2. £ADOWANIE DANYCH PACJENTA
Public Sub LoadPatientData(PatientName As String)
    Dim ws As Worksheet: Set ws = Sheets("Baza_danych_medycznych")
    Dim foundRng As Range
    
    Set foundRng = ws.Columns("A").Find(PatientName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRng Is Nothing Then
        rIndex = foundRng.Row
        txtWywiad.Text = ws.Cells(rIndex, 2).Value
        txtBadanie.Text = ws.Cells(rIndex, 3).Value
        txtRozpoznanie.Text = ws.Cells(rIndex, 4).Value
        txtZalecenia.Text = ws.Cells(rIndex, 5).Value
        RefreshVisitList ws.Cells(rIndex, 6).Value, txtFiltrDaty.Text
    Else
        Application.EnableEvents = False
        rIndex = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(rIndex, 1).Value = PatientName
        Application.EnableEvents = True
        ClearFields
    End If
End Sub

' 3. ODŒWIE¯ANIE LISTY
Sub RefreshVisitList(rawData As String, Optional FilterText As String = "")
    Dim v() As String, p() As String, temp As String
    Dim i As Long, j As Long
    Dim d As String, n As String, o As String, St As String
    Dim wpisListy As String
    
    lstWizyty.Clear
    If Trim(rawData) = "" Then Exit Sub
    
    ' Rozbijamy dane
    v = Split(rawData, "||")
    
    ' SORTOWANIE B¥BELKOWE (Wed³ug DATY, potem Numeru)
    For i = LBound(v) To UBound(v) - 1
        For j = i + 1 To UBound(v)
            If Len(v(i)) > 5 And Len(v(j)) > 5 Then
                Dim p1() As String: p1 = Split(v(i), "|")
                Dim p2() As String: p2 = Split(v(j), "|")
                
                ' Porównanie: najpierw data (p(0)), potem numer (p(1))
                If p1(0) > p2(0) Then
                    temp = v(i): v(i) = v(j): v(j) = temp
                ElseIf p1(0) = p2(0) Then
                    If CLng(p1(1)) > CLng(p2(1)) Then
                        temp = v(i): v(i) = v(j): v(j) = temp
                    End If
                End If
            End If
        Next j
    Next i
    
    ' WYŒWIETLANIE NA LIŒCIE
    For i = 0 To UBound(v)
        If Len(Trim(v(i))) > 5 Then
            p = Split(v(i), "|")
            d = p(0): n = p(1)
            o = IIf(UBound(p) >= 2, IIf(Trim(p(2)) = "", "Brak opisu", p(2)), "Brak opisu")
            
            St = "Nieop³acono"
            For j = 0 To UBound(p)
                If p(j) = "Op³acono" Or p(j) = "Nieop³acono" Then St = p(j): Exit For
            Next j
            
            wpisListy = d & " (Nr: " & n & ") - " & o & " [" & St & "]"
            
            If FilterText = "" Or InStr(1, wpisListy, FilterText, vbTextCompare) > 0 Then
                lstWizyty.AddItem wpisListy
            End If
        End If
    Next i
End Sub

' 4. ZAPIS DANYCH OGÓLNYCH
Private Sub btnZapisz_Click()
    Dim ws As Worksheet: Set ws = Sheets("Baza_danych_medycznych")
    Application.EnableEvents = False
    ws.Cells(rIndex, 2).Value = txtWywiad.Text
    ws.Cells(rIndex, 3).Value = txtBadanie.Text
    ws.Cells(rIndex, 4).Value = txtRozpoznanie.Text
    ws.Cells(rIndex, 5).Value = txtZalecenia.Text
    Application.EnableEvents = True
    MsgBox "Dane pacjenta zosta³y zapisane!", vbInformation
End Sub

' 5. DODAWANIE NOWEJ WIZYTY
Private Sub btnDodajWizyte_Click()
    Dim vDate As String, vNum As String, vDesc As String, vStatus As String
    Dim record As String, currentContent As String
    Dim ws As Worksheet: Set ws = Sheets("Baza_danych_medycznych")
    Dim maxNr As Long, i As Long
    Dim v() As String, p() As String
    
    currentContent = Trim(ws.Cells(rIndex, 6).Value)
    maxNr = 0
    If currentContent <> "" Then
        v = Split(currentContent, "||")
        For i = 0 To UBound(v)
            If Len(v(i)) > 5 Then
                p = Split(v(i), "|")
                If UBound(p) >= 1 Then
                    If IsNumeric(p(1)) Then If CLng(p(1)) > maxNr Then maxNr = CLng(p(1))
                End If
            End If
        Next i
    End If
    
    vDate = InputBox("Data (RRRR-MM-DD):", "Nowa Wizyta", Format(Date, "yyyy-mm-dd"))
    If vDate = "" Then Exit Sub
    
    ' Walidacja daty
    If Not IsDate(vDate) Then
        MsgBox "B³êdny format daty! Wpisz np. 2026-04-23", vbCritical
        Exit Sub
    End If
    vDate = Format(CDate(vDate), "yyyy-mm-dd")
    
    vNum = InputBox("Numer wizyty:", "Nowa Wizyta", maxNr + 1)
    If vNum = "" Then Exit Sub
    
    vDesc = InputBox("Opis wizyty:", "Nowa Wizyta")
    If Trim(vDesc) = "" Then vDesc = "Brak opisu"
    
    vStatus = IIf(cmbStatusPlatnosci.Value = "", "Nieop³acono", cmbStatusPlatnosci.Value)
    record = vDate & "|" & vNum & "|" & vDesc & "|" & vStatus
    
    Application.EnableEvents = False
    If currentContent = "" Then
        ws.Cells(rIndex, 6).Value = record
    Else
        ws.Cells(rIndex, 6).Value = currentContent & "||" & record
    End If
    Application.EnableEvents = True
    
    RefreshVisitList ws.Cells(rIndex, 6).Value, txtFiltrDaty.Text
    cmbStatusPlatnosci.Value = "Nieop³acono"
End Sub
' 6. AKTUALIZACJA WYBRANEJ WIZYTY
Private Sub btnAktualizujWizyte_Click()
    Dim ws As Worksheet: Set ws = Sheets("Baza_danych_medycznych")
    Dim v() As String, p() As String
    Dim i As Long, pelnyString As String, nData As String, nOpis As String, nNum As String
    Dim wybranyTekst As String, nrDoAktualizacji As String
    
    If lstWizyty.ListIndex = -1 Then
        MsgBox "Wybierz wizytê z listy!", vbExclamation
        Exit Sub
    End If
    
    ' Wyci¹gamy numer wizyty z zaznaczonego wiersza na liœcie
    wybranyTekst = lstWizyty.List(lstWizyty.ListIndex)
    nrDoAktualizacji = Mid(wybranyTekst, InStr(wybranyTekst, "Nr: ") + 4)
    nrDoAktualizacji = Left(nrDoAktualizacji, InStr(nrDoAktualizacji, ")") - 1)
    
    pelnyString = ws.Cells(rIndex, 6).Value
    v = Split(pelnyString, "||")
    
    For i = 0 To UBound(v)
        p = Split(v(i), "|")
        If p(1) = nrDoAktualizacji Then
            nData = InputBox("Popraw datê (RRRR-MM-DD):", "Aktualizacja", p(0))
            If nData = "" Then Exit Sub
            If Not IsDate(nData) Then MsgBox "B³êdna data!": Exit Sub
            
            nNum = InputBox("Popraw numer wizyty:", "Aktualizacja", p(1))
            If nNum = "" Then Exit Sub
            
            nOpis = InputBox("Popraw opis:", "Aktualizacja", IIf(UBound(p) >= 2, p(2), ""))
            
            v(i) = Format(CDate(nData), "yyyy-mm-dd") & "|" & nNum & "|" & nOpis & "|" & cmbStatusPlatnosci.Value
            Exit For
        End If
    Next i
    
    pelnyString = ""
    For i = 0 To UBound(v)
        If Len(Trim(v(i))) > 0 Then
            pelnyString = pelnyString & IIf(pelnyString = "", "", "||") & v(i)
        End If
    Next i
    
    Application.EnableEvents = False
    ws.Cells(rIndex, 6).Value = pelnyString
    Application.EnableEvents = True
    
    RefreshVisitList pelnyString, txtFiltrDaty.Text
    MsgBox "Wizyta zosta³a zaktualizowana!", vbInformation
End Sub

' 7. USUWANIE WIZYTY
Private Sub btnUsunWizyte_Click()
    Dim ws As Worksheet: Set ws = Sheets("Baza_danych_medycznych")
    Dim v() As String, p() As String, i As Long, nowyString As String
    Dim wybranyTekst As String, nrDoUsuniecia As String
    
    If lstWizyty.ListIndex = -1 Then
        MsgBox "Wybierz wizytê do usuniêcia!", vbExclamation
        Exit Sub
    End If
    
    ' Pobranie numeru z tekstu na liœcie
    wybranyTekst = lstWizyty.List(lstWizyty.ListIndex)
    nrDoUsuniecia = Mid(wybranyTekst, InStr(wybranyTekst, "Nr: ") + 4)
    nrDoUsuniecia = Left(nrDoUsuniecia, InStr(nrDoUsuniecia, ")") - 1)
    
    If MsgBox("Czy na pewno chcesz usun¹æ wizytê nr " & nrDoUsuniecia & "?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    v = Split(ws.Cells(rIndex, 6).Value, "||")
    nowyString = ""
    
    For i = 0 To UBound(v)
        If Len(Trim(v(i))) > 0 Then
            p = Split(v(i), "|")
            If p(1) <> nrDoUsuniecia Then
                nowyString = nowyString & IIf(nowyString = "", "", "||") & v(i)
            End If
        End If
    Next i
    
    Application.EnableEvents = False
    ws.Cells(rIndex, 6).Value = nowyString
    Application.EnableEvents = True
    
    RefreshVisitList nowyString, txtFiltrDaty.Text
End Sub

' 8. OTWIERANIE PLIKÓW
Private Sub btnOtworzPliki_Click()
    Dim sciezka As String, nazwa As String
    nazwa = Cells(ActiveCell.Row, "D").Value & "_" & Cells(ActiveCell.Row, "E").Value
    sciezka = ThisWorkbook.Path & S & "Dokumentacja_Pacjentow" & S & nazwa
    
    If Dir(sciezka, vbDirectory) <> "" Then
        #If Mac Then
            Shell "open " & """" & sciezka & """"
        #Else
            Shell "explorer.exe " & """" & sciezka & """", vbNormalFocus
        #End If
    Else
        If MsgBox("Utworzyæ folder?", vbYesNo + vbQuestion) = vbYes Then
            Dim glowny As String: glowny = ThisWorkbook.Path & S & "Dokumentacja_Pacjentow"
            If Dir(glowny, vbDirectory) = "" Then MkDir glowny
            MkDir sciezka
            #If Mac Then
                Shell "open " & """" & sciezka & """"
            #Else
                Shell "explorer.exe " & """" & sciezka & """", vbNormalFocus
            #End If
        End If
    End If
End Sub

' 9. DRUKOWANIE
Private Sub btnDrukujOgolne_Click()
    Dim wsS As Worksheet: Set wsS = Sheets("Szablon_Druk")
    Dim pacjent As String, folderGlowny As String, folderPacjenta As String, nazwaBezSpacji As String
    Dim r As Long, sciezka As String
    
    r = ActiveCell.Row
    pacjent = Sheets("Pacjenci").Cells(r, "D").Value & " " & Sheets("Pacjenci").Cells(r, "E").Value
    If Trim(pacjent) = "" Or pacjent = " " Then Exit Sub

    wsS.Range("C2").Value = pacjent
    wsS.Range("C4").Value = Format(Date, "dd.mm.yyyy")
    wsS.Range("B7").Value = txtWywiad.Text
    wsS.Range("B14").Value = txtBadanie.Text
    wsS.Range("B22").Value = txtRozpoznanie.Text
    wsS.Range("B30").Value = txtZalecenia.Text

    folderGlowny = ThisWorkbook.Path & S & "Wydruki"
    nazwaBezSpacji = Replace(pacjent, " ", "_")
    folderPacjenta = folderGlowny & S & nazwaBezSpacji
    
    If Dir(folderGlowny, vbDirectory) = "" Then MkDir folderGlowny
    If Dir(folderPacjenta, vbDirectory) = "" Then MkDir folderPacjenta

    sciezka = folderPacjenta & S & "Karta_Ogolna_" & nazwaBezSpacji & ".pdf"
    wsS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sciezka, OpenAfterPublish:=True
End Sub

' 10. DRUKOWANIE WIZYT
Private Sub btnDrukujWizyte_Click()
    Dim wsS As Worksheet, wsB As Worksheet, wsP As Worksheet
    Dim pacjent As String, folderPacjenta As String, nazwaPliku As String
    Dim rawData As String, v() As String, p() As String
    Dim wybraneNumery As String, numeryTab() As String
    Dim i As Long, j As Long, startRow As Long, currentRow As Long
    Dim rP As Long, licznikWydruku As Integer
    
    On Error Resume Next
    Set wsS = Sheets("Szablon_Druk _2")
    Set wsB = Sheets("Baza_danych_medycznych")
    Set wsP = Sheets("Pacjenci")
    On Error GoTo 0
    
    If wsS Is Nothing Then Exit Sub

    rP = ActiveCell.Row
    pacjent = wsP.Cells(rP, "D").Value & " " & wsP.Cells(rP, "E").Value
    wybraneNumery = InputBox("Wpisz numery wizyt po przecinku (np. 1, 4, 7):", "Wybór wizyt")
    If wybraneNumery = "" Then Exit Sub
    
    numeryTab = Split(wybraneNumery, ",")
    wsS.Range("C2").Value = pacjent
    wsS.Range("C4").Value = Format(Date, "dd.mm.yyyy")
    
    Dim lastRowS As Long: lastRowS = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
    If lastRowS >= 9 Then wsS.Range("A9:B" & lastRowS + 5).Value = ""

    rawData = wsB.Cells(rIndex, 6).Value
    v = Split(rawData, "||")
    startRow = 9: licznikWydruku = 0
    
    For i = 0 To UBound(numeryTab)
        Dim szukanyNr As String: szukanyNr = Trim(numeryTab(i))
        For j = 0 To UBound(v)
            If Len(v(j)) > 5 Then
                p = Split(v(j), "|")
                If p(1) = szukanyNr Then
                    currentRow = startRow + (licznikWydruku * 3)
                    wsS.Cells(currentRow, "A").Value = "Data: " & p(0)
                    wsS.Cells(currentRow, "B").Value = "Wizyta nr " & p(1) & ": " & p(2) & vbCrLf & "[" & p(3) & "]"
                    licznikWydruku = licznikWydruku + 1
                    Exit For
                End If
            End If
        Next j
    Next i

    folderPacjenta = ThisWorkbook.Path & S & "Wydruki" & S & Replace(pacjent, " ", "_")
    If Dir(ThisWorkbook.Path & S & "Wydruki", vbDirectory) = "" Then MkDir ThisWorkbook.Path & S & "Wydruki"
    If Dir(folderPacjenta, vbDirectory) = "" Then MkDir folderPacjenta
    
    nazwaPliku = folderPacjenta & S & "Wybrane_Wizyty_" & Replace(pacjent, " ", "_") & "_" & Format(Date, "yyyyMMdd") & ".pdf"
    wsS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nazwaPliku, OpenAfterPublish:=True
End Sub

' 11. POMOCNICZE
Sub ClearFields()
    txtWywiad.Text = "": txtBadanie.Text = "": txtRozpoznanie.Text = "": txtZalecenia.Text = ""
    txtFiltrDaty.Text = "": lstWizyty.Clear: cmbStatusPlatnosci.Value = "Nieop³acono"
End Sub

Private Sub lstWizyty_Click()
    Dim ws As Worksheet: Set ws = Sheets("Baza_danych_medycznych")
    Dim v() As String, p() As String, i As Integer
    If lstWizyty.ListIndex <> -1 Then
        v = Split(ws.Cells(rIndex, 6).Value, "||")
        p = Split(v(lstWizyty.ListIndex), "|")
        cmbStatusPlatnosci.Value = "Nieop³acono"
        For i = 0 To UBound(p)
            If p(i) = "Op³acono" Or p(i) = "Nieop³acono" Then cmbStatusPlatnosci.Value = p(i): Exit For
        Next i
    End If
End Sub

Private Sub txtFiltrDaty_Change()
    If rIndex > 0 Then RefreshVisitList Sheets("Baza_danych_medycznych").Cells(rIndex, 6).Value, txtFiltrDaty.Text
End Sub
