Attribute VB_Name = "Cursor"
Option Compare Database

#If VBA7 Or Win64 Then
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
#Else
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
#End If


' lpCursorName-Konstanten für LoadCursor
Public Const IDC_APPSTARTING As Long = 32650   ' Blauer, sich drehender Kreis
Public Const IDC_ARROW       As Long = 32512   ' Standardpfeil
Public Const IDC_CROSS       As Long = 32515   ' Kreuz, Fadenkreuz
Public Const IDC_HAND        As Long = 32649   ' Hand
Public Const IDC_HELP        As Long = 32651   ' Pfeil mit Fragezeichen
Public Const IDC_IBEAM       As Long = 32513   ' Textcursor (Einfügemarke)
Public Const IDC_ICON        As Long = 32641   ' -
Public Const IDC_NO          As Long = 32648   ' roter Kreis mit Schrägstrich ("Nicht ablegen")
Public Const IDC_SiZE        As Long = 32640   ' -
Public Const IDC_SIZEALL     As Long = 32646   ' Kreuz mit Pfeilen
Public Const IDC_SIZENESW    As Long = 32643   ' Pfeil von Nordost nach Südwest
Public Const IDC_SIZENS      As Long = 32645   ' Pfeil von Nord nach Süd
Public Const IDC_SIZENWSE    As Long = 32642   ' Pfeil von Nordwest nach Südost
Public Const IDC_SIZEWE      As Long = 32644   ' Pfeil von West nach Ost
Public Const IDC_UPARROW     As Long = 32516   ' Pfeil nach oben
Public Const IDC_WAIT        As Long = 32514   ' Blauer, sich drehender Kreis
                                               ' vormals: Sanduhr, "Warten")

Public Sub CursorWechseln(CursorTyp As String)

    Dim lHandle As Long

    lHandle = LoadCursor(0, CursorTyp)
    If (lHandle > 0) Then SetCursor lHandle

    'Aufruf über z.B.  Call CursorWechseln(IDC_HAND)

End Sub
