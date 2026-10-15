VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim pacjentFull As String
    
    ' 1. Reaguje tylko na kolumny D lub E od wiersza 7 w dó³
    If Not Intersect(Target, Range("D7:E1000")) Is Nothing Then
        If Target.Value = "" Then Exit Sub
        
        ' 2. Wy³¹czamy zdarzenia
        Application.EnableEvents = False
        
        Cancel = True ' Blokuje mruganie kursora w komórce
        
        ' Pobieranie danych
        pacjentFull = Cells(Target.Row, "D").Value & " " & Cells(Target.Row, "E").Value
        
        ' 3. Otwieranie formularza
        With frmKartaPacjenta
            .LoadPatientData pacjentFull
            .Show
        End With
        
        ' 4. W³¹czamy zdarzenia z powrotem po zamkniêciu formularza
        Application.EnableEvents = True
    End If
End Sub

' BLOKADA RÊCZNEJ EDYCJI
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Jeœli zmiana nast¹pi³a w kolumnach D lub E
    If Not Intersect(Target, Range("D7:E1000")) Is Nothing Then
        
        ' Sprawdzamy, czy makra s¹ wolne" (EnableEvents = True)
        ' Jeœli dwuklik lub makro przycisku dzia³aj¹, to poni¿szy kod siê nie odpali
        If Application.EnableEvents = True Then
            
            On Error Resume Next
            Application.EnableEvents = False ' Blokada pêtli
            
            Application.Undo ' Cofa rêczne wpisanie/skasowanie
            
            MsgBox "Nie mo¿esz edytowaæ tabeli rêcznie!" & vbCrLf & _
                   "U¿yj przycisków 'Dodaj dane' lub 'Usuñ dane'.", vbCritical, "Blokada edycji"
            
            Application.EnableEvents = True
        End If
    End If
End Sub

