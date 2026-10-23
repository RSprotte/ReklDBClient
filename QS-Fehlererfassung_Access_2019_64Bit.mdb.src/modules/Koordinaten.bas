Attribute VB_Name = "Koordinaten"
Option Compare Database
Option Explicit

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function GetWindowRect Lib "user32" _
            (ByVal hwnd As Long, lpRect As RECT) As Long
#Else
    Private Declare Function GetWindowRect Lib "user32" _
            (ByVal hwnd As Long, lpRect As RECT) As Long
#End If


Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type


Public Function setFktKoordinaten(ByRef frm As Form) As String
   'Menüleiste - Extras - Verweise - "Microsoft DAO3.6 Object Library" - Häkchen setzen
    Dim r As RECT
    Dim RHF As RECT
    Dim rs As DAO.Recordset
    Dim ldb As DAO.Database
    Dim bWrongFrmSizeOrPos As Boolean
    Dim lLeft As Long
    Dim lTop As Long
    Dim lHigh As Long
    Dim lWidth As Long

    bWrongFrmSizeOrPos = False
    GetWindowRect Forms("TB:autostart").hwnd, RHF 'Basis
    GetWindowRect frm.hwnd, r 'Dynamisches fenster
    If r.Left * 15 > 31680 Then
        r.Left = 31680 / 15
        bWrongFrmSizeOrPos = True
    End If
    'Größe/Breite anhand neuer Position korregieren
    lLeft = r.Left - RHF.Left 'Relativ zum Hauptfenster

    If r.Top * 15 > 31680 Then
        r.Top = 31680 / 15
        bWrongFrmSizeOrPos = True
    End If
    'Größe/Höhe anhand neuer Position korregieren
    lTop = r.Top - RHF.Top

    If (r.Right - r.Left) * 15 > 31680 Then
        lWidth = 31680 / 15
        bWrongFrmSizeOrPos = True
    Else
        lWidth = r.Right - r.Left
    End If

    If (r.Bottom - r.Top) * 15 > 31680 Then
        lHigh = 31680 / 15
        bWrongFrmSizeOrPos = True
    Else
        lHigh = r.Bottom - r.Top
    End If


    If bWrongFrmSizeOrPos Then
        MsgBox getTC_FCFromCode("TC0304")
        'MsgBox ("Das Fenster wurde auf die max. Größe sowie Position korregiert!")  'TC0304
    End If
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * " & _
                                      "FROM   Tab_Koordinaten " & _
                                      "WHERE  Formname='" & frm.Name & "'", dbOpenDynaset)
    If rs.RecordCount Then
        ldb.Execute ("Update Tab_Koordinaten set Koordinaten = '" & lLeft * 15 & "," & lTop * 15 & "," & _
                         lWidth * 15 & "," & lHigh * 15 & "' where Formname = '" & frm.Name & "';")
    Else
        ldb.Execute ("Insert into Tab_Koordinaten (Formname, Koordinaten) SELECT '" & frm.Name & "' AS Ausdr1, '" & lLeft * 15 & "," & lTop * 15 & "," & _
                         lWidth * 15 & "," & lHigh * 15 & "' AS Ausdr2;")
    End If

    rs.Close
    Set rs = Nothing
'    ldb.Close
     ' Set ldb = openAccDBConnection("closeDB")

    setFktKoordinaten = lLeft * 15 & "," & lTop * 15 & "," & _
                         lWidth * 15 & "," & lHigh * 15
    'MsgBox ("Der Dialog wurden mit den Paramter Links: " & lLeft * 15 & " Oben: " & lTop * 15 & " Breite: " & lWidth * 15 & " Höhe: " & lHigh * 15 & " festgelegt!") 'Lokalisierung

End Function

Public Function initFktKoordinaten(ByRef frm As Form) As String
   'Menüleiste - Extras - Verweise - "Microsoft DAO3.6 Object Library" - Häkchen setzen
    Dim r As RECT
    Dim RHF As RECT
    Dim rs As DAO.Recordset
    Dim ldb As DAO.Database
    Dim bWrongFrmSizeOrPos As Boolean
    Dim lLeft As Long
    Dim lTop As Long
    Dim lHigh As Long
    Dim lWidth As Long
    Dim sKor As String
    Dim sKorSplit() As String
    Dim varKoord As Variant

    bWrongFrmSizeOrPos = False
    GetWindowRect Forms("TB:autostart").hwnd, RHF 'Basis

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * " & _
                                      "FROM   Tab_Koordinaten " & _
                                      "WHERE  Formname='" & frm.Name & "'", dbOpenDynaset)
    If rs.RecordCount Then
        sKor = rs("Koordinaten")
        sKorSplit = Split(sKor, ",")
        ldb.Execute ("Update Tab_Koordinaten set Koordinaten = '0,0," & _
                         sKorSplit(2) & "," & sKorSplit(3) & "' where Formname = '" & frm.Name & "';")
        sKorSplit = getFktKoordinatenRelHF(frm)
        If sKorSplit(3) > 0 And sKorSplit(2) > 0 Then
        'Fenster setzen relativ zum Hauptfenster TB:autostart
            DoCmd.MoveSize sKorSplit(0), sKorSplit(1), sKorSplit(2), sKorSplit(3)
        End If
    End If


    rs.Close
    Set rs = Nothing
'    ldb.Close
     ' Set ldb = openAccDBConnection("closeDB")

    initFktKoordinaten = "0,0," & _
                         sKorSplit(2) & "," & sKorSplit(3)
    'MsgBox ("Der Dialog wurden mit den Paramter Links: 0 Oben: 0 Breite: " & sKorSplit(2) & " Höhe: " & sKorSplit(3) & " festgelegt!") 'Lokalisierung

End Function

Public Function getFktKoordinaten(ByRef frm As Form) As String
   'Menüleiste - Extras - Verweise - "Microsoft DAO3.6 Object Library" - Häkchen setzen
    Dim r As RECT
    Dim rs As DAO.Recordset
    Dim ldb As DAO.Database

    GetWindowRect frm.hwnd, r
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * " & _
                                      "FROM   Tab_Koordinaten " & _
                                      "WHERE  Formname='" & frm.Name & "'", dbOpenDynaset)
    If rs.RecordCount Then
        getFktKoordinaten = rs("Koordinaten")
    Else
        getFktKoordinaten = ""
        'MsgBox getTC_FCFromCode("TC0305", ldb)
        'MsgBox ("Keine Koordinaten vorhanden! Koordinaten durch Positionierung und Größenanpassung noch festlegen!")    'TC0305
    End If

    rs.Close
    Set rs = Nothing
'    ldb.Close
     ' Set ldb = openAccDBConnection("closeDB")


End Function

Public Function getFktKoordinatenRelHF(ByRef frm As Form) As Variant
'   'Menüleiste - Extras - Verweise - "Microsoft DAO3.6 Object Library" - Häkchen setzen
'    Dim r As RECT
'    Dim RHF As RECT
'    Dim rs As DAO.Recordset
'    Dim ldb As DAO.Database
'    Dim varKoord As Variant
'
'    GetWindowRect Forms("TB:autostart").hWnd, RHF
'    GetWindowRect frm.hWnd, r
'    Set ldb = openAccDBConnection("openDB")
'    Set rs = ldb.OpenRecordset("SELECT * " & _
'                                      "FROM   Tab_Koordinaten " & _
'                                      "WHERE  Formname='" & frm.Name & "'", dbOpenDynaset)
'
'
'    If rs.RecordCount Then
'        varKoord = Split(Nz(rs("Koordinaten"), "4000,4000,4000,2000"), ",")
'        If varKoord(0) + (RHF.Left * 15) > 31680 Then
'            varKoord(0) = 31680
'        Else
'            varKoord(0) = varKoord(0) + RHF.Left * 15
'        End If
'        If varKoord(1) + (RHF.Top * 15) > 31680 Then
'            varKoord(1) = 31680
'        Else
'            varKoord(1) = varKoord(1) + RHF.Top * 15
'        End If
'
'        getFktKoordinatenRelHF = varKoord
'    Else
'        varKoord = Split(Nz(Null, "0,0,0,0"), ",")
'        getFktKoordinatenRelHF = varKoord
'        'MsgBox getTC_FCFromCode("TC0305", ldb)
'        'MsgBox ("Keine Koordinaten vorhanden! Koordinaten durch Positionierung und Größenanpassung noch festlegen!")    'TC0306
'    End If
'
'    rs.Close
'    Set rs = Nothing
''    ldb.Close
'     ' Set ldb = openAccDBConnection("closeDB")


End Function
