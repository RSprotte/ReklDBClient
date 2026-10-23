Attribute VB_Name = "FensterHWnd"
Option Compare Database
Option Explicit

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                        ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                        ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If




Public blFenster As Boolean


Private Const WM_DESTROY = &H2
Private Const WM_CLOSE = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SW_HIDE = 0       ' Fenster ausblenden
Private Const SW_NORMAL = 1     ' Fenster wiederherstellen bzw. einblenden
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOW = 5

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Public Function GetFormPosition(frm As Form, sngLeft As Single, sngTop As Single, sngWidth As Single, _
                                sngHeight As Single, sngRight As Single)

'Funktion zur Ermittlung Koordinaten und Größe des übergebenen Formular
'Nullpunkt ist hierbei Hauptbildschirm Links Oben

    Dim rectForm As RECT
    Dim rectParent As RECT

    'Ermittelt die Position des Formulars auf dem Desktop
    GetWindowRect frm.hwnd, rectForm
'    GetWindowRect Application.hWndAccessApp, rectParent

    'Abstand zum außersten Bildschirmrand Rechts
    sngRight = rectForm.Right

    'Abstand zum außersten Bildschirmrand Links
    sngLeft = rectForm.Left '- rectParent.Left

    'Abstand zum außersten Bildschirmrand Oben
    sngTop = rectForm.Top '- rectParent.Top

    'Errechnung der Breite des Formular
    sngWidth = rectForm.Right - rectForm.Left  '- rectParent.Left

    'Errechnung der Höhe des Formular
    sngHeight = rectForm.Bottom - rectForm.Top  '- rectParent.Top

End Function

Public Sub FormKoorSave(ByRef frm As Form)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim s_left As Single
    Dim s_top As Single
    Dim s_width As Single
    Dim s_heigt As Single
    Dim s_right As Single
    Dim strKoor As String

    'Ermitteln der Koordinaten
    GetFormPosition frm, s_left, s_top, s_width, s_heigt, s_right
    'String zur Eingabe der Koordinaten und Größe in Tabelle
    strKoor = s_left & "/" & s_top & "/" & s_width & "/" & s_heigt & "/" & s_right

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * FROM   Tab_Koordinaten WHERE  Formname='" & frm.Name & "'", dbOpenDynaset)

    'Überprüfen ob ein Tabellen-Eintrag für das entsprechende Formular in der Koordinaten-Tabelle vorhanden ist
    'wenn ja wird dieser editiert wenn nicht wird ein neuer Datensatz angelegt
    If rs.RecordCount Then
        rs.MoveFirst
        rs.Edit
        rs("Koordinaten") = strKoor
        rs.Update
    Else
        ldb.Execute ("Insert into Tab_Koordinaten ([Formname],[Koordinaten]) values ('" & frm.Name & "','" & strKoor & "');")
    End If


    Set rs = Nothing
     Set ldb = openAccDBConnection("closeDB")

End Sub

Public Sub FormOpenKoor(frm As Form)

    Dim varKoord As Variant
    Dim strDaten As String
    Dim s_left As Single
    Dim nResult As Long


    strDaten = Nz(DLookup("Koordinaten", "Tab_Koordinaten", "[Formname] = '" & frm.Name & "'"))
    varKoord = Split(strDaten, "/")
    'varKoord(0) = von links ,varKoord(1) = von oben ,varKoord(2) = breite ,varKoord(3) = höhe ,varKoord(4) = von rechts

    If strDaten = "" Then
        'Posizioniert das Fenster mit festgelegten Koordinaten und Größe wenn noch kein Eintrag in der Koordinaten-Tabelle erfolgt ist
        SetWindowPos frm.hwnd, 0, 100, 100, 600, 600, SWP_NOZORDER

    Else
        'GetSystemMetrics(78) ermittelt den Maximalen Anzeigebereich (Breite) aller Bildschirme von Rechts nach Links
        'GetSystemMetrics(76) X Wert ins Minus Bildschirm
        'GetSystemMetrics(0) X Wert Hauptbildschirm

        'Ermittlung des Maximale X+ und X- Koordinate Anzeigebereich (Breite)
        If varKoord(4) >= GetSystemMetrics(78) + GetSystemMetrics(76) Then
            s_left = GetSystemMetrics(0) - varKoord(2)
            ElseIf varKoord(0) <= GetSystemMetrics(76) Then
                s_left = GetSystemMetrics(76)
        Else
            s_left = varKoord(0)
        End If

        'Posizioniert das Fenster mit den gespeicherten Koordinaten und Größe aus dem Tabellen-Eintrag
        SetWindowPos frm.hwnd, 0, s_left, varKoord(1), varKoord(2), varKoord(3), SWP_NOZORDER
    End If

End Sub

Public Sub AccessFenster()

    Dim hWindow As Long
    Dim nResult As Long
    Dim i As Integer
    Dim strfrm As String

    hWindow = Application.hWndAccessApp
'    nResult = ShowWindow(hWindow, 10)

    If blFenster Then 'Ausgeblendet
        'einblenden
'        CommandBars.ActiveMenuBar.Enabled = True
'        CommandBars("Database").Visible = True
        DoCmd.SelectObject acTable, , True
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
        Call Application.SetOption("Show Status Bar", True)
'        nResult = ShowWindow(hWindow, SW_SHOW) 'SW_SHOW
        blFenster = False
    Else
        'ausblenden
'        CommandBars.ActiveMenuBar.Enabled = False
'        CommandBars("Database").Visible = False
        DoCmd.SelectObject acTable, , True
        RunCommand acCmdWindowHide
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
        Call Application.SetOption("Show Status Bar", False)
'        nResult = ShowWindow(hWindow, SW_HIDE) 'SW_HIDE
        blFenster = True
    End If

'    For i = 0 To (Forms().Count - 1)
'        strfrm = Forms(i).Name
'        nResult = ShowWindow(Forms(strfrm).hwnd, SW_NORMAL)
'    Next i


'    For i = 0 To Forms.Count - 1
'        strfrm = Forms(i).Caption
'        If LCase$(Forms(i).Caption) = LCase$(strfrm) Then
'          nResult = ShowWindow(Forms(i).hWnd, SW_NORMAL)
'        End If
'    Next i

    Forms("TB:autostart").SetFocus

End Sub

Public Function OpenWin(ByRef frm As Form, Optional lEdit As Boolean = True) As Form
'Blendet das Formular ein und wird zum Bearbeiten (Optional) freigegeben
    Dim nResult As Long

    frm.AllowEdits = lEdit
    nResult = ShowWindow(frm.hwnd, SW_NORMAL)
    Set OpenWin = frm

End Function

Public Sub closeWinHandle(ByRef frm As Form)

'    Dim lngReturn As Long
'
'    lngReturn = CloseWindow(frm.hwnd)
'    frm.Visible = False

    'Neuer Code
    FormKoorSave frm
    SendMessage frm.hwnd, WM_CLOSE, 0, 0

End Sub

Public Sub FormEinUndAusBlenden(ByRef frm As Form, Optional lSichtbar As Boolean = True)

    Dim nResult As Long

    If lSichtbar = True Then
        nResult = ShowWindow(frm.hwnd, SW_NORMAL)
    Else
        nResult = ShowWindow(frm.hwnd, SW_HIDE)
    End If

End Sub
Public Sub ReportEinUndAusBlenden(ByRef fRp As Report, Optional lSichtbar As Boolean = True)

    Dim nResult As Long

    If lSichtbar = True Then
        nResult = ShowWindow(fRp.hwnd, SW_NORMAL)
    Else
        nResult = ShowWindow(fRp.hwnd, SW_HIDE)
    End If

End Sub

Public Sub FormOpenKoorMousPointer(ByRef frm As Form, ByVal fxpos As Long, ByVal fypos As Long, ByVal fBreite As Long, ByVal fHoehe As Long, Optional fMittig As Boolean = True)

    Call FormEinUndAusBlenden(frm)
    'Ermittlung des Maximale X+ und X- Koordinate Anzeigebereich (Breite)
    If fxpos + fBreite > GetSystemMetrics(78) + GetSystemMetrics(76) Then
        fxpos = (GetSystemMetrics(78) + GetSystemMetrics(76)) - fBreite
    Else
        If fMittig = True Then fxpos = fxpos - (fBreite / 2)
    End If

    If fMittig = True Then fypos = fypos - (fHoehe / 2)

    'Posizioniert das Fenster mit den gespeicherten Koordinaten und Größe aus dem Tabellen-Eintrag
    SetWindowPos frm.hwnd, 0, fxpos, fypos, fBreite, fHoehe, SWP_NOZORDER

End Sub
