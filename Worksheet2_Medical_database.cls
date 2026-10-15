VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Jeœli zdarzenia s¹ w³¹czone, znaczy ¿e to rêczna zmiana
    If Application.EnableEvents = True Then
        On Error Resume Next
        Application.EnableEvents = False
        Application.Undo
        MsgBox "Baza danych medycznych jest chroniona. Edycja tylko przez formularz!", vbCritical
        Application.EnableEvents = True
    End If
End Sub
