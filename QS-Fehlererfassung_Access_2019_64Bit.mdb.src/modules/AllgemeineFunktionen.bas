Attribute VB_Name = "AllgemeineFunktionen"

Option Compare Database
Option Explicit

Public pl_RueckliefTemp As String

Public strValueInputBox As String
Public bolStop As Boolean
'Public get_pKuerzel() As String
Public pl_KuerzelMWSB As Long 'Verantwortlicher SB Mahnung
Public pl_KuerzelMWPR As Long 'Verantwortlicher SB Wirksamkeitsprüfung
Public pi_Kuerzel As Long
Public p_ReklNr As String
Public p_KommentarKopieren As Boolean
Public p_VorgangStatus As String
Public p_ReklMemoQuelle As String


Public DialogFehlerbeschreibungGrob As String
Public FehlerbeschreibungGrobID As Integer
Public PFehlerbeschreibungGrob As String

Public DialogTaetigkeit As String
Public TaetigkeitID As Integer
Public PTaetigkeit As String
Public DialogDefTaetigkeit As String
Public DefTaetigkeitID As Integer
Public PDefTaetigkeit As String
Public DialogFehlerbildAnalyse As String
Public FehlerbildAnalyseID As Integer

Public psDialogKunde As String
Public psDialogAnsprechpartnerKunde As String
'### Änderung
Public plDialogSelectedKundenID As Long
Public plDialogSelectedAPID As Long

Public psMassnahmeBearbeitung As String

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
            Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare Function GetUserName Lib "advapi32.dll" _
                         Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                               nSize As Long) As Long

    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
#End If


#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Private MousePosition As POINTAPI
Private Type POINTAPI
    X As Long
    Y As Long
End Type


Const GWL_EXSTYLE As Long = -20
Const WS_EX_DLGMODALFRAME As Long = &H1

Type OPENFILENAME

    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long

End Type


Const OFN_FILEMUSTEXIST = &H1000
Const OFN_READONLY = &H1
Const OFN_HIDEREADONLY = &H4

'RSP - 20131110 - Neu
Public Const Massnahme = "SELECT Maßnahmen.*, getBearbeitungsortMaßnahme(Maßnahmen.MaßnahmeID) AS BearbeitungsortAll, Sachbearbeiter.Kurzzeichen " & _
                  "FROM Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]" & _
                  "ORDER BY Maßnahmen.MaßnahmeID;"
Public Const MassnahmeReport = "SELECT Maßnahmen.*, getBearbeitungsortMaßnahme(Maßnahmen.MaßnahmeID) AS BearbeitungsortAll, Sachbearbeiter.Kurzzeichen, getFehlerklasseMaßnahme(Maßnahmen.MaßnahmeID) AS FehlerklasseAll " & _
                  "FROM Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]" & _
                  "ORDER BY Maßnahmen.MaßnahmeID;"
Public Const MassnahmeMitKlasse = "SELECT Maßnahmen.*, MaßnahmenFehlerarten.*, [Kategorie-BearbeitungsOrtMaßnahme].BearbeitungsortMaßnahme, Sachbearbeiter.Kurzzeichen " & _
                           "FROM ((Maßnahmen LEFT JOIN [Kategorie-BearbeitungsOrtMaßnahme] ON Maßnahmen.BearbeitungsOrt = [Kategorie-BearbeitungsOrtMaßnahme].[MB-ID]) " & _
                           "LEFT JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID) LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID] " & _
                           "ORDER BY Maßnahmen.MaßnahmeID;"
Public Const MassnahmeMitKlasseReport = "SELECT Maßnahmen.*, MaßnahmenFehlerarten.*, getBearbeitungsortMaßnahme(Maßnahmen.MaßnahmeID) AS BearbeitungsortAll, Sachbearbeiter.Kurzzeichen, getFehlerklasseMaßnahme(Maßnahmen.MaßnahmeID) AS FehlerklasseAll " & _
                           "FROM (Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
                           "ORDER BY Maßnahmen.MaßnahmeID;"


'Public Const cReklamationsVerzeichnis = "F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\"
'Public Const cPERechercheHistorie = "F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\Fehler-Recherche PE Historie\"
''Public Const cReklamationsVerzeichnis = "C:\Users\ralf sprotte\Documents\Fehlerberichte_internerGebrauch\"

Public Const maxTime = 10    ' in seconds
Public Const sleepTime = 250 ' in milliseconds
Public sLQArtikelTemp As String
Public lLQArtikelTemp As Long
Public sLQBGTemp As String
Public lLQBGTemp As Long
Public sKundenTemp As String
Public lKundenTemp As Long
Public sSelectDatumForm As String
Public sCurrentFehlererfassung As String
Public sReklNrReferenzPos As String
Public sReklNrReferenz As String

Public gLoggingClient As String
Public gLoggingClientsAll As String

'Mail Absender Eingabefelder
Public sMailTo As String
Public sMailCC As String
Public sMailBCC As String

'Reklamationbericht zum Lieferant
Public sBestellnummerLQ As String
Public sLieferscheinLieferant As String
Public sReparaturdatum As String

'Smartpointer
Type sptrClass
    sptr As Object
    sptrAnzRef As Integer
End Type

Public sptrReklVorgang As sptrClass

Public bKeyStrg As Boolean
Public sCurrArtikelSource As String

Public gVerlängertBis As Date

'## 20181217-ROS Änderung
Public ps_MailAn As String
Public ps_MailCC As String
Public ps_MailBCC As String

'#### Rückgabeparameter forms
Public psgetValueFromInputBox As String

#If VBA7 Then
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As LongPtr, ByVal dwExtraInfo As LongPtr)
#Else
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

' Tasten-Konstanten für die offizielle Copilot-Tastenkombination
Private Const VK_LWIN As Byte = &H5B       ' Windows-Taste
Private Const VK_SHIFT As Byte = &H10      ' Shift-Taste
Private Const VK_F23 As Byte = &H86        ' F23-Taste
Private Const KEYEVENTF_KEYUP As Long = &H2

Public Sub Copilot_Shortcut_Druecken()
    ' 1. WIN und SHIFT herunterdrücken
    keybd_event VK_LWIN, 0, 0, 0
    keybd_event VK_SHIFT, 0, 0, 0

    ' 2. F23 drücken und loslassen
    keybd_event VK_F23, 0, 0, 0
    keybd_event VK_F23, 0, KEYEVENTF_KEYUP, 0

    ' 3. WIN und SHIFT loslassen
    keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
End Sub

Public Function getSystemInfo() As Boolean  '20210420-ROS

    #If VBA7 Or Win64 Then
        getSystemInfo = True
    #Else
        getSystemInfo = False
    #End If

End Function


Public Sub ListeSortieren(Liste As Object)
    Dim ZeileA As Integer
    Dim ZeileB As Integer
    Dim SortierteListe As String, TEMP
    Dim ListenEinträge()
    ReDim ListenEinträge(0 To Liste.ListCount - 1)
    Dim iZwischenspeicher As String

    ' In Datenfeld einlesen und sortieren
    For ZeileA = 0 To Liste.ListCount - 1
        ListenEinträge(ZeileA) = Liste.Column(0, ZeileA)
        For ZeileB = ZeileA To 1 Step -1
            If ListenEinträge(ZeileB) < ListenEinträge(ZeileB - 1) Then
                TEMP = ListenEinträge(ZeileB)
                ListenEinträge(ZeileB) = ListenEinträge(ZeileB - 1)
                ListenEinträge(ZeileB - 1) = TEMP
            End If
        Next ZeileB
    Next ZeileA

    ' An Liste zurückgeben
    SortierteListe = vbNullString
    For ZeileA = 0 To Liste.ListCount - 1
        If InStr(1, ListenEinträge(ZeileA), ";") = 0 Then
            If iZwischenspeicher <> ListenEinträge(ZeileA) Then
                SortierteListe = SortierteListe & ";" & ListenEinträge(ZeileA)
                iZwischenspeicher = ListenEinträge(ZeileA)
            End If
        Else
            SortierteListe = SortierteListe & ";""" & ListenEinträge(ZeileA) & """"
        End If
    Next ZeileA

    Liste.RowSource = MID(SortierteListe, 2)
End Sub

Public Sub set_Flag_Leitersatz_For_all_Artikel_in_Rekl()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Artikel-all_avista_out];")
    rs.MoveLast
    rs.MoveFirst
    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset

    While Not rs.EOF

        'Recordset-Instanz
        Set rsAvista = adbprov.getRecordsetAvista("Select tm_artikel.ast_nummer, tm_artikel.ast_bezeichnung AS artikel_bezeichnung from tm_artikel WHERE ast_nummer = '" & rs("LQ Art-Nr") & "' AND tm_artikel.ast_id IN ( SELECT xm_artikel.ast_id FROM DBA.xm_artikel  WHERE xm_artikel.ast_artikel_typ = 1)")
'        If rs("LQ Art-Nr") = "1101376" Then
'            Debug.Print rs("LQ Art-Nr")
'        End If
        If rsAvista.RecordCount Then
            rs.Edit
            rs("AvistaArtikelTyp") = 1
            rs.Update
        End If
        rs.MoveNext
    Wend
    Set rs = Nothing
    Set rsAvista = Nothing
    Set adbprov = Nothing
    adbprov.closeAvistaConnection

End Sub


Public Function get_p_ReklNr() As String   '20160729-ROS

    While StrPtr(p_ReklNr) = 0 Or p_ReklNr = ""
'        p_ReklNr = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
        p_ReklNr = getValueFromInputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
        'get_get_p_ReklNr ()l () = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")
    Wend
    get_p_ReklNr = p_ReklNr

End Function
Public Function get_pKuerzel(Optional ByVal lKz As String, Optional ByVal blKz As Boolean = False) As String  '20160729-ROS

    If IsNull(TempVars(cVAR_Kuerzel)) Or (TempVars(cVAR_Kuerzel)) = "" Then
        While IsNull(TempVars(cVAR_Kuerzel)) Or (TempVars(cVAR_Kuerzel)) = ""
'            TempVars.Add cVAR_Kuerzel, InputBox(getTC_FCFromCode("TC3188"), getTC_FCFromCode("TC0211"))
            If blKz = True Then
                TempVars.Add cVAR_Kuerzel, lKz
            Else
                TempVars.Add cVAR_Kuerzel, getValueFromInputBox(getTC_FCFromCode("TC3188"), getTC_FCFromCode("TC0211"))
            End If
                If (TempVars(cVAR_Kuerzel)) = "" Then get_pKuerzel
                If IsNumeric(TempVars(cVAR_Kuerzel)) Then     '20170131-ROS
'                        MsgBox ("Nur Buchstabeneingabe möglich")
                        MsgBox getTC_FCFromCode("TC3187"), vbInformation, getTC_FCFromCode("TC3188")
                        TempVars.Remove (cVAR_Kuerzel)
                        get_pKuerzel
                End If

        Wend
    End If

    get_pKuerzel = StrConv(TempVars(cVAR_Kuerzel), 1)     '20121121-RSP Änderung


End Function

Public Function DateiOeffnen(Optional Titel, Optional Filter, Optional DefExtension, Optional AktDir) As String
On Error GoTo err_DateiOeffnen

    Dim strDateiname As String
    Dim strDlgTitel As String
    Dim strFilter As String
    Dim strDefExtension As String
    Dim strAktDir As String
    Dim strNull As String
'    Dim OpenDlg As OPENFILENAME
    Dim OpenDlg As Office.FileDialog
    Set OpenDlg = Application.FileDialog(msoFileDialogFilePicker)

    strNull = Chr$(0)
    strDateiname = String$(512, 0)

    If IsMissing(Titel) Then
        strDlgTitel = "Open File" & strNull  'getTC_FCFromCode("TC0436") & strNull
    Else
        strDlgTitel = Titel & strNull
    End If

    If IsMissing(Filter) Then
        strFilter = "All files" & strNull & "*.*" & strNull & strNull   'getTC_FCFromCode("TC0437") & strNull & "*.*" & strNull & strNull
    Else
        strFilter = Filter & strNull
    End If

    If IsMissing(DefExtension) Then
        strDefExtension = strNull
    Else
        strDefExtension = DefExtension & strNull
    End If

    If IsMissing(AktDir) Then
        strAktDir = CurDir$ & strNull
    Else
        strAktDir = AktDir & strNull
    End If

    With OpenDlg
        .Title = "Datenbank Auswählen"                    'Fenstertitel
        .AllowMultiSelect = False                 'Nur eine Datei auswählbar
        .ButtonName = "Auswählen"                 'Button Beschriftung
        .Filters.Clear                            'erst alle Filter löschen
        .InitialFileName = "C:\"                  'Startverzeichnis
        .Show
    End With

    If OpenDlg.SelectedItems.Count > 0 Then
        DateiOeffnen = OpenDlg.SelectedItems(1)
    Else
        DateiOeffnen = ""
    End If

'    With OpenDlg
'        .lStructSize = Len(OpenDlg)
'        '.hwndOwner = Screen.ActiveForm.Hwnd
'        .lpstrFilter = strFilter
'        .nFilterIndex = 1
'        .lpstrFile = strDateiname
'        .nMaxFile = Len(strDateiname)
'        .lpstrInitialDir = strAktDir
'        .lpstrTitle = strDlgTitel
'        .Flags = OFN_FILEMUSTEXIST Or OFN_READONLY
'        .lpstrDefExt = strDefExtension
'        If GetOpenFileName(OpenDlg) <> 0 Then
'            DateiOeffnen = Left$(.lpstrFile, InStr(.lpstrFile, strNull) - 1)
'        Else
'            DateiOeffnen = ""
'        End If
'    End With

    Exit Function

err_DateiOeffnen:
    MsgBox "error: " & Err & " - " & Error
End Function
Public Function DateiMultiOeffnen(Optional Titel, Optional Filter, Optional DefExtension, Optional AktDir) As Object

    Dim sFiles As Object
    Dim sFile2 As Variant
    Dim strNull As String

    strNull = Chr$(0)


    Set sFiles = Application.FileDialog(3)
    sFiles.AllowMultiSelect = True
    If IsMissing(AktDir) Then
        sFiles.InitialFileName = CurDir$ & strNull
    Else
        sFiles.InitialFileName = AktDir & strNull
    End If

    sFiles.Show

    If True Then
        Set DateiMultiOeffnen = sFiles
    Else
        Set DateiMultiOeffnen = Null
    End If
End Function


Function IsLoaded(ByVal strFormname As String) As Integer

'Returns a 0 if form is not open or a -1 if Open

    If SysCmd(acSysCmdGetObjectState, acForm, strFormname) <> 0 Then

        If Forms(strFormname).CurrentView <> 0 Then

            IsLoaded = True

        End If

    End If

End Function



Function InsertNewEinkaufDS(ByVal fi_reklid As Long) As Boolean

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

 Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
Set rs = ldb.OpenRecordset("Select * from ROS_Daten_Einkauf where [ID-ReklNr] = " & fi_reklid & ";")

If rs.RecordCount = 0 Then
    'Vorgang wird neu angelegt
    ldb.Execute ("Insert into ROS_Daten_Einkauf ([ID-ReklNr]) values (" & fi_reklid & ");")
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")
InsertNewEinkaufDS = True

End Function

Function InsertNewReklMemoDS(ByVal fi_reklid As Long, Optional ByVal fZusatztext As String = "") As Boolean

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Stundensatz As String
Dim sOut As String

Stundensatz = getTC_FCFromCode("DP0008") '20180815-ROS variabler Stundensatz eingefügt konfigurierbar über RSP-Lokalisierung

 Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & fi_reklid & ";")

If rs.RecordCount = 0 Then
    'Vorgang wird neu angelegt
    If DLookup("[Intern]", "Mastertab", "ID=" & fi_reklid) Then
        ldb.Execute ("Insert into ReklMemo (ID, [Aktueller Eintrag], VorgangstypID,[AktuellerStundensatz]) values (" & fi_reklid & ", '" & _
                    format(Now()) & " - " & get_pKuerzel() & ": Vorgang angelegt.', 1," & Stundensatz & ");") '20180815-ROS variabler Stundensatz eingefügt konfigurierbar über RSP-Lokalisierung
    Else
        ldb.Execute ("Insert into ReklMemo (ID, [Aktueller Eintrag],[AktuellerStundensatz]) values (" & fi_reklid & ", '" & _
                    format(Now()) & " - " & get_pKuerzel() & ": Vorgang angelegt.', " & Stundensatz & ");") '20180815-ROS variabler Stundensatz eingefügt konfigurierbar über RSP-Lokalisierung

    End If
Else

    sOut = format(Now()) & " - " & get_pKuerzel() & ":"
    If Not fZusatztext = "" Then
        sOut = sOut & Chr(13) & Chr(10) & fZusatztext & "!"
    End If
'##################################################################     '20190830-ROS_1
'wird nicht mehr benötigt VorgangstypID als Standart 1(Meldung) und Sundensatz 45€ in DB gesetzt
'    If DLookup("[Intern]", "Mastertab", "ID=" & fi_reklid) Then
'        If rs("VorgangstypID") = 0 Then
'            rs.Edit
'            rs("VorgangstypID") = 1
'            rs("AktuellerStundensatz") = Stundensatz        '20190829-ROS_1
'            rs.Update
'        End If
'    End If
'###################################################################
    ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = '" & sOut & _
                "' where ID = " & fi_reklid & ";")
    p_KommentarKopieren = True
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")
InsertNewReklMemoDS = True

End Function
Public Function InsertMemoDS(ByVal fi_reklid As Long, fMemo As String) As Boolean

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim sOut As String

 Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & fi_reklid & ";")
'Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & Forms("Fehlererfassung - Register").Controls("ID") & ";")

If rs.RecordCount > 0 Then

    sOut = format(Now()) & " - " & get_pKuerzel() & " - Reklamation " & fi_reklid & " in Bearbeitung."
    ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = '" & rs("Aktueller Eintrag") & fMemo & _
                "' where ID = " & fi_reklid & ";")
End If
InsertMemoDS = True


Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Function


Public Function getTageQuery(ID As Long) As Integer

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim i As Integer

 Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem

Set rs = ldb.OpenRecordset("SELECT RSP_Ausw_Basisdaten_1.ID, Format([RSP_Ausw_Basisdaten_1].[Auftragsdatum],'yyyy.mm') AS ADATUM, RSP_Ausw_Basisdaten_1.[WE-Datum], " & _
                           "RSP_Ausw_Basisdaten_1.[Vorgang - Pos20], RSP_Ausw_Basisdaten_1.[Vorgang - Pos30], RSP_Ausw_Basisdaten_1.[Vorgang - Pos40], " & _
                           "RSP_Ausw_Basisdaten_1.[Vorgang - Pos50], RSP_Ausw_Basisdaten_1.[Vorgang - Pos60], RSP_Ausw_Basisdaten_1.[Vorgang - Pos70], " & _
                           "RSP_Ausw_Basisdaten_1.[Vorgang - Pos80], Mastertab.[Datum - Pos20], Mastertab.[Datum - Pos30], Mastertab.[Datum - Pos40], " & _
                           "Mastertab.[Datum - Pos50], Mastertab.[Datum - Pos60], Mastertab.[Datum - Pos70], Mastertab.[Datum - Pos80] " & _
                           "FROM Mastertab INNER JOIN RSP_Ausw_Basisdaten_1 ON Mastertab.ID = RSP_Ausw_Basisdaten_1.ID " & _
                           "Where [RSP_Ausw_Basisdaten_1].[ID] = " & ID & ";")
For i = 2 To 8
    If rs.Fields("Vorgang - Pos" & i & "0").Value = "26" Then
        'MsgBox ("Datum: " & Me.Recordset.Fields(i + 7).Value)
        If (IsNull(rs.Fields("Datum - Pos" & i & "0").Value) Or rs.Fields("Datum - Pos" & i & "0").Value = "") Then
            getTageQuery = 0
        Else
            getTageQuery = fncArbeitstage(rs.Fields("WE-Datum").Value, rs.Fields("Datum - Pos" & i & "0").Value)
        End If
        'getDatum = Me.Controls("Datum - Pos" & i & "0").Value
        Exit For
    End If
Next i

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Function




Public Function fncArbeitstage( _
    datStart As Date, _
    datEnde As Date) _
 As Integer

    'Deklaration
    Dim intWochen As Integer
    Dim intTage As Integer
    Dim varDatum As Variant

    'Anzahl ganze Wochen zwischen den Daten
    intWochen = DateDiff("w", datStart, datEnde)
    If intWochen <= 1 Then
        intWochen = 0
    End If

    'Anzahl Arbeitstage zwischen den Daten (5 Tage pro Woche)
    intTage = intWochen * 5
    'Datum, das nach der letzten vollen Woche im Datumsbereich liegt
    varDatum = DateAdd("ww", intWochen, datStart)
    'Für restliche Tage:
    'Daten, die nicht auf ein Wochenende fallen dazuzählen
    Do While varDatum < datEnde
        If Weekday(varDatum) <> vbSunday And Weekday(varDatum) <> vbSaturday Then
            intTage = intTage + 1
        End If
        'Nächstes zuprüfendes Datum
        varDatum = DateAdd("d", 1, varDatum)
    Loop

    'Rückgabe
    fncArbeitstage = intTage

 End Function



Public Function get_LQ_Einheit(LQEinheit_ID As Long) As String

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

 Set ldb = openAccDBConnection("openDB")

Set rs = ldb.OpenRecordset("select * from LQEinheit where ID=" & LQEinheit_ID & ";")
If rs.RecordCount > 0 Then
    get_LQ_Einheit = rs("LQEinheit").Value
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Function



Public Function LinkTable(strDatabaseSource As String, strTableSource As String, strTableDestination As String)

'    Dim dbSource As DAO.Database
'
'    Dim dbTarget As DAO.Database
'
'    Dim tdf As DAO.TableDef
'
'    Dim dbDestination As DAO.Database
'    Dim i As Integer
'
'    On Error GoTo LinkTable_Err
'
'    Set dbSource = DBEngine.Workspaces(0).OpenDatabase("C:\Users\ralf sprotte\Desktop\QS-Fehlererfassung_Daten.mdb")
'
'    Set dbDestination = CurrentDb
'
'    For i = 0 To dbDestination.TableDefs.Count - 1
'        dbDestination.TableDefs(i).Delete
'    Next i
'    Set tdf = dbDestination.CreateTableDef(strTableDestination)
'
'    tdf.Connect = ";DATABASE=" & strDatabaseSource
'
'    tdf.SourceTableName = strTableSource
'
'    dbDestination.TableDefs.Append tdf
'
'    LinkTable = True
'
'LinkTable_Exit:
'
'    dbSource.Close
'
'    Set dbSource = Nothing
'
'    Set dbDestination = Nothing
'
'    Set tdf = Nothing
'
'    Exit Function
'
'LinkTable_Err:
'
'    LinkTable = False
'
'    Resume LinkTable_Exit

End Function


Public Function getFehlerklasseMaßnahme(fMID As Long) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    getFehlerklasseMaßnahme = ""
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select MaßnahmenFehlerarten.Fehlerklasse FROM MaßnahmenFehlerarten WHERE (((MaßnahmenFehlerarten.MaßnahmeID)=" & fMID & "));")
    While Not rs.EOF
        getFehlerklasseMaßnahme = getFehlerklasseMaßnahme & rs("Fehlerklasse") & Chr(13) & Chr(10)
        rs.MoveNext
    Wend
    If Len(getFehlerklasseMaßnahme) > 0 Then
        getFehlerklasseMaßnahme = Left(getFehlerklasseMaßnahme, Len(getFehlerklasseMaßnahme) - 2)
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")


End Function
Public Function getBearbeitungsortMaßnahme(fMID As Long) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    getBearbeitungsortMaßnahme = ""
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT ([LQEinheit] & '->' & [LQBereich]) AS Bearbeitungsort FROM MaßnahmenFehlerortGruppe INNER JOIN " & _
                              "(LQEinheit INNER JOIN LQBereich ON LQEinheit.ID = LQBereich.[ID-LQEinheit]) ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID " & _
                              "WHERE (((MaßnahmenFehlerortGruppe.MaßnahmeID)=" & fMID & "));")
    While Not rs.EOF
        getBearbeitungsortMaßnahme = getBearbeitungsortMaßnahme & rs("Bearbeitungsort") & Chr(13) & Chr(10)
        rs.MoveNext
    Wend
    If Len(getBearbeitungsortMaßnahme) > 0 Then
        getBearbeitungsortMaßnahme = Left(getBearbeitungsortMaßnahme, Len(getBearbeitungsortMaßnahme) - 2)
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")


End Function



Public Function getClientVersion() As String
On Error GoTo Err_getClientVersion
'RS - 2012-06-11 Verionausgabe eingebunden
Dim ldb As DAO.Database
Dim rs, rsClientLang As DAO.Recordset


'RS - 2012-06-11 Verionausgabe eingebunden
 Set ldb = openAccDBConnection("openDB")
Set rs = ldb.OpenRecordset("SELECT VersionierungClient.* FROM VersionierungClient " & _
                          "ORDER BY VersionierungClient.Version DESC;")
Set rsClientLang = ldb.OpenRecordset("SELECT KonfigurationReklDBLocal.Value FROM KonfigurationReklDBLocal " & _
                          "WHERE KonfigurationReklDBLocal.Konfiguration = 'Location';")

rs.MoveFirst
rsClientLang.MoveFirst
getClientVersion = Application.CurrentProject.FullName & " - Version: " & rs("Version") & "  " & rsClientLang("Value") & " - " & rs("Datum")

Set rs = Nothing
Set rsClientLang = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End_getClientVersion:
    Exit Function

Err_getClientVersion:

    If Err = -2147467259 Then
        Resume Next
    End If

End Function

Public Function getServerLocation() As String
On Error GoTo err_getServerLocation
'RSP - 20131110 Erweiterung
'Anzeige der zentralen Datenbank auf der Startmaske mit dem der Client verbunden ist
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim t As TableDef

Set ldb = openAccDBConnection("openDB", "New")

Set t = ldb.TableDefs("Versionierung")
Set rs = ldb.OpenRecordset("SELECT Versionierung.* FROM Versionierung " & _
                          "ORDER BY Versionierung.Version DESC;")

rs.MoveFirst

getServerLocation = Right(t.Connect, Len(t.Connect) - InStr(1, t.Connect, "=")) & " - " & getTC_FCFromCode("TC0438") & " " & rs("Version") & "  -  " & rs("Datum")

Set rs = Nothing

end_getServerLocation:
Set t = Nothing
' Set ldb = openAccDBConnection("closeDB")
Exit Function

err_getServerLocation:
    If Err = 3024 Then
        MsgBox (Error)
        GoTo end_getServerLocation
    End If
    'Resume Next
End Function



Function delCRLRandMinus(fs As String) As String

    Dim l_spos As Integer
    Dim l_cpos As Integer   'position before search
    Dim r_String As String

    l_spos = 1
    l_cpos = 1
    r_String = ""
    'While l_spos = InStr(l_spos, fs, Chr(13), vbBinaryCompare) > 0
    l_spos = InStr(l_spos, fs, Chr(13))
    While (l_spos > 0)
'        Debug.Print MID(fs, l_spos - 1, 1)
        If MID(fs, l_spos - 1, 1) = "-" Then
            l_spos = l_spos - 1
            r_String = r_String & MID(fs, l_cpos, l_spos - l_cpos) & " "
            l_spos = l_spos + 3
        Else
'            Debug.Print MID(fs, l_cpos, l_spos - l_cpos)
            r_String = r_String & MID(fs, l_cpos, l_spos - l_cpos) & " "
            l_spos = l_spos + 2
        End If
        l_cpos = l_spos
        l_spos = InStr(l_spos, fs, Chr(13))
    Wend
'    Debug.Print MID(fs, l_cpos, (Len(fs) - l_cpos) + 1)
    r_String = r_String & MID(fs, l_cpos, (Len(fs) - l_cpos) + 1)

    'Ende text einfügen
    delCRLRandMinus = r_String

End Function


Function fOSUserName() As String
     Dim lngLen As Long, lngX As Long
     Dim strUserName As String

     strUserName = String$(254, 0)
     lngLen = 255
     lngX = GetUserName(strUserName, lngLen)
     If lngX <> 0 Then
        fOSUserName = Left$(strUserName, lngLen - 1)
       Else
        fOSUserName = ""
     End If
End Function


Public Function getReklVerzeichnis(ByVal fReklNr As Long, ByVal fJahr As String, Optional fMeldungsnummer As String = "") As String

    Dim cReklamationsVerzeichnis As String   '20160304-ROS
    cReklamationsVerzeichnis = getTC_FCFromCode("DP0004")

    If fMeldungsnummer = "" Then
        If Dir(cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\", vbDirectory) <> "" Then
            getReklVerzeichnis = cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\"
        Else
            getReklVerzeichnis = ""
        End If
    Else
        If Dir(cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "_" & fMeldungsnummer & "\", vbDirectory) <> "" Then
            getReklVerzeichnis = cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "_" & fMeldungsnummer & "\"
        Else
            getReklVerzeichnis = ""
        End If
    End If

End Function
Public Function setReklVerzeichnis(ByVal fReklNr As Long, ByVal fJahr As String, Optional fMeldungsnummer As String = "") As String

    Dim cReklamationsVerzeichnis As String   '20160304-ROS
    cReklamationsVerzeichnis = getTC_FCFromCode("DP0004")
'20160322-RSP
On Error GoTo err_setReklVerzeichnis

    'Ist Verzeichnis mit ReklNr vorhanden
    If getReklVerzeichnis(fReklNr, fJahr) = "" Then
        If ((fMeldungsnummer = "") Or (IsNull(fMeldungsnummer))) Then
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\")
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0463") & "\")      '20161020-ROS Ordner Bilder
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0464") & "\")      '20161020-ROS Ordner Stellungnhame
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0465") & "\")      '20161020-ROS Ordner E-Mail
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0483") & "\")      '20170428-ROS Ordner Kundendoku
                setReklVerzeichnis = cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\"
       Else
            If getReklVerzeichnis(fReklNr, fJahr, fMeldungsnummer) = "" Then
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "_" & fMeldungsnummer & "\")
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\")
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0463") & "\")      '20161020-ROS Ordner Bilder
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0464") & "\")      '20161020-ROS Ordner Stellungnhame
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0465") & "\")      '20161020-ROS Ordner E-Mail
                MkDir (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\" & getTC_FCFromCode("TC0483") & "\")      '20170428-ROS Ordner Kundendoku
            End If
                setReklVerzeichnis = cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "_" & fMeldungsnummer & "\"
        End If
    Else
        'Vorhanden
        'Wenn eine Meldungsnummer  mittlerweile eingetragen worden ist wird das Verzeichnis
        'mit der Meldungsnummer umbenannt
        If (fMeldungsnummer = "" Or IsNull(fMeldungsnummer)) Then
            setReklVerzeichnis = cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\"
        Else
            Name (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "\") As _
                 (cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "_" & fMeldungsnummer & "\")
            setReklVerzeichnis = cReklamationsVerzeichnis & fJahr & "\" & fReklNr & "_" & fMeldungsnummer & "\"
        End If
    End If
    Exit Function
err_setReklVerzeichnis:

    If Err = 76 Then 'Pfad nicht gefunden (...Netzwerk nicht vorhanden!!
        MsgBox getTC_FCFromCode("TC0473")
        'MsgBox "Pfad kann nicht angelegt werden! Netzwerk vorhanden?"
        setReklVerzeichnis = ""
        Exit Function 'Beenden da Verzeichnis nicht angelegt werden kann
    ElseIf Err = 58 Then 'Verzeichnis existiert bereits   20161104-ROS
            'MsgBox ("Vorgangsordner doppelt vorhanden!" & vbCrLf & "Ordner ohne Kd_Meldungsnummer löschen!" & vbCrLf & "Vor dem Löschen, Inhalt in anderen Ordner kopieren!")
            MsgBox (getTC_FCFromCode("TC0474") & vbCrLf & getTC_FCFromCode("TC0475") & vbCrLf & getTC_FCFromCode("TC0476"))
        Resume Next
    End If


End Function

Public Function getTRArtNrAvista(NewData As String) As Long
'20140222-RSP: Ermittelt alle Artikel in Avista anhand der TRUMPF Artikelnummer

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim ldb As DAO.Database
    Dim rs2 As DAO.Recordset


    getTRArtNrAvista = 0
    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    Set rs = adbprov.getRecordsetAvista(AvistaKndArtNrSQL_Select & NewData & "' " & AvistaArtikelSQL_Order)
    If rs.RecordCount Then
        'If DCount("[Art-ID]", "Artikel-all_avista_out", "[Kd Mat-Nr] = '" & NewData & "'") = 0 Then
        Set ldb = openAccDBConnection("openDB")
        Set rs2 = ldb.OpenRecordset("Select * from [Artikel-all_avista_out] where [Kd Mat-Nr] = '" & NewData & "';")
        If rs2.RecordCount = 0 Then
            rs2.AddNew
            rs2![LQ Art-Nr] = rs("ast_nummer")
            rs2![Variante] = rs("avar_nummer")
            rs2![Kd Mat-Nr] = rs("avkd_artnr_kd")
        Else
            rs2.Edit
        End If
        rs2![Bezeichnung 1] = rs("ast_bezeichnung")
        rs2![Hersteller-Teilenummer] = rs("avar_herst_nummer")
        getTRArtNrAvista = rs2![Artikel-ID]
        rs2.Update
        rs2.Close
        Set rs2 = Nothing
'        ldb.Close
         ' Set ldb = openAccDBConnection("closeDB")
    End If
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

ende:
    Exit Function

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")     'TC0292
    Resume ende

End Function

Public Sub getLQArtNrAvista(AID As String)

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim Response

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Fehlerausgabe unterdrücken
    Response = acDataErrAdded
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & AID & "' " & AvistaArtikelSQL_Order)
    If rs.RecordCount Then
        DoCmd.OpenForm "Artikel-avista_outTemp", acFormDS, , , , acDialog
        Set Forms("Artikel-avista_outTemp").Recordset = rs
        Forms("Artikel-avista_outTemp").Requery
    End If
    Response = acDataErrContinue
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")     'TC0292
    Resume ende
End Sub

Public Function fgetLQArtNrAvista(AID As String) As ADODB.Recordset

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim Response

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Fehlerausgabe unterdrücken
    Response = acDataErrAdded
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    Set rs = adbprov.getRecordsetAvista(AvistaArtikelSQL_Select & AID & "' " & AvistaArtikelSQL_Order)
    If rs.RecordCount Then
        'Set getLQArtNrAvista = rs
'        DoCmd.OpenForm "Artikel-avista_outTemp", acFormDS, , , , acDialog
'        Set Forms("Artikel-avista_outTemp").Recordset = rs
'        Forms("Artikel-avista_outTemp").Requery
    End If
'    Response = acDataErrContinue
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

ende:
    Exit Function
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")     'TC0292
    Resume ende
End Function


Public Sub NewReklamationE(fReason, f_ReklNr, Optional fState As String = "", Optional fb_SB As Boolean = True, Optional fb_LockDS As Boolean = True)

Dim sPfadDatei As String
Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean
Dim ReklBearbeiten As Boolean
Dim lb_reload As Boolean        'Wird benötigt um ein Formular nochmals zu laden (Event Load) obwohl beim DoCmd.open das Formular schon geöffnet ist.
                                'In dem Fall wird das Event Load nicht aufgerufen
Dim lf_FE_Form As [Form_Fehlererfassung - Register]
Dim Test As Boolean
lb_reload = False


Dim cReklamationsVerzeichnis As String   '20160304-ROS
cReklamationsVerzeichnis = getTC_FCFromCode("DP0004")

'Wenn Grund = "Bearbeiten, dann muss eine ReklNummer übergeben werden. Ist ads nicht der Fall wird der Dialog zur Eiongabe einer ReklNr erneut gestartet
If fReason = "Bearbeiten" Then
    If f_ReklNr = "" Then
        f_ReklNr = get_p_ReklNr()
    Else
        p_ReklNr = f_ReklNr
    End If
Else
    p_ReklNr = ""
End If


NewRekl = False

If p_ReklNr = "" And get_ServiceZuordnung = False Then
'    MsgBox ("Externe Reklamationen können nur von einem Service-Mitarbeiter angelegt werden!")
    MsgBox getTC_FCFromCode("TC3189"), vbInformation, getTC_FCFromCode("TC3190")
    Exit Sub
End If

If p_ReklNr = "" And fReason <> "Bearbeiten" Then
    'ermitteln letzte ReklNr
    p_ReklNr = getLastDSMastertab()
    If get_p_ReklNr() > 0 Then
        'ReklNr erhöhen = Neuer Vorgang
        p_ReklNr = get_p_ReklNr() + 1
        'Öffnen der Fehlererfassung mit neuer Rekl-Nr
        ' -> Notwendig damit die Reklamtionsnummer sofort generiert wird und für weiter Zwecke genutzt werden kann
        If fState = "" Then
'            If (setNewReklDS(get_p_ReklNr(), Format(Date, "dd.mm.yy"), "Bearbeitung: Service", "berechtigt", "extern") = 0) Then Exit Sub
            If (setNewReklDS(get_p_ReklNr(), format(Date, "dd\/mm\/yyyy"), "Bearbeitung: Service", "berechtigt", "extern") = 0) Then Exit Sub
        Else
'            If (setNewReklDS(get_p_ReklNr(), Format(Date, "dd.mm.yy"), fState, "berechtigt", "extern") = 0) Then Exit Sub
            If (setNewReklDS(get_p_ReklNr(), format(Date, "dd\/mm\/yyyy"), fState, "berechtigt", "extern") = 0) Then Exit Sub
        End If
        '20151002-RSP Filter direkt setzen
'        If fb_LockDS Then DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab.ID] = " & get_p_ReklNr(), acFormEdit
        'Das aktuelle Datum wird als Auftragsdatum eingetragen
        'Ermittlung Zuständiger Service Mitarbeiter (Erfasser - Extern)
        Set ldb = openAccDBConnection("openDB")
        If fb_SB Then
            Set rs = ldb.OpenRecordset("Select * from Bearbeiter where [Kurzzeichen] = '" & get_pKuerzel() & "';", dbOpenSnapshot)
'            Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Kurzzeichen] = '" & get_pKuerzel() & "';")
            If rs.RecordCount Then
                ldb.Execute ("Update Mastertab set [Zuständiger Service-MA] = '" & rs("Name").Value & "' where [ID] = " & get_p_ReklNr() & ";")
                ldb.Execute ("Update Mastertab set [Verursacher] = 'LQ' where [ID] = " & get_p_ReklNr() & ";")
                ldb.Execute ("Update Mastertab set [Menge betroffen - Kunde] = 1 where [ID] = " & get_p_ReklNr() & ";")
'                Forms("Fehlererfassung - Register").Requery
                NewPosAufwand get_p_ReklNr(), rs("Name").Value
            Else
                MsgBox Replace(getTC_FCFromCode("TC0434"), "<xxx>", get_pKuerzel()), vbOKOnly, getTC_FCFromCode("TC0407")
            End If
        End If
        'Template Fehlergrundbeschreibung in Bearbeiter Daten 2
'        ldb.Execute ("Update Mastertab set [Wie Warum passierte der Fehler] = 'Fehlergrund:" & _
'            Chr(13) & Chr(10) & _
'            Chr(13) & Chr(10) & _
'            "Qualifizierung: berechtigt" & _
'            Chr(13) & Chr(10) & "' where [ID] = " & get_p_ReklNr() & ";")

        'Änderungen für eine Einheitliche Bearbeitung der Fehlerbeschreibung in Reklamation 20260209-ROS
        ldb.Execute ("Update Mastertab set [Wie Warum passierte der Fehler] = '" & _
            getTC_FCFromCode("TC4259") & _
            Chr(13) & Chr(10) & _
            Chr(13) & Chr(10) & _
            getTC_FCFromCode("TC4260") & _
            Chr(13) & Chr(10) & _
            Chr(13) & Chr(10) & _
            getTC_FCFromCode("TC4261") & _
            Chr(13) & Chr(10) & _
            Chr(13) & Chr(10) & _
            getTC_FCFromCode("TC4262") & _
            Chr(13) & Chr(10) & _
            Chr(13) & Chr(10) & _
            getTC_FCFromCode("TC4263") & _
            Chr(13) & Chr(10) & "' where [ID] = " & get_p_ReklNr() & ";")

        'Anlegen des Datensatz zum Sperren der Bearbeitung durch einen Zweiten Benutzer
        If fb_LockDS Then lockDataset (get_p_ReklNr())
        'ReklBearbeiten = True
        '
        'InsertNewReklMemoDS (get_get_p_ReklNr ()l ())
        'Beim Öffnen so setzen das der generierte Kommentar nicht kopiert wird. Kopie macht erst Sinn wenn Änderungen
        'am Datensatz bzw. direkt im Kommentarfeld vollzogen wurden. In dem Fall wird die Variable wie folgt geändert:
        'Bei einer Änderung im Steuerelement wird die Variable auf True gesetzt (
        p_KommentarKopieren = False

        NewRekl = True
        'Ordner für die Dokumentenablage anlegen - Ordnername = ReklNr
        If Dir(getTC_FCFromCode("DP0004") & format(Date, "yyyy") & "\" & p_ReklNr & "\", vbDirectory) = "" Then
            sPfadDatei = setReklVerzeichnis(p_ReklNr, format(Date, "yyyy"))
            'MkDir (getTC_FCFromCode("DP0004") & Format(Date, "yyyy") & "\" & p_ReklNr & "\")
        End If
        If fb_LockDS Then
            If IsLoaded("Fehlererfassung - Register") Then lb_reload = True
            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
            If lb_reload Then Call Forms("Fehlererfassung - Register").ReLoadForm
        End If
    End If

Else
    'gibt es die eingegeben Reklamationsnummer?

    Dim strKrit As String
    strKrit = "ID = " & get_p_ReklNr()
    If DCount("ID", "Mastertab", strKrit) = 0 Then
        MsgBox Replace(getTC_FCFromCode("TC0440"), "<xxx>", get_p_ReklNr())    '("Reklamation mit der Nummer: " & get_get_p_ReklNr ()l () & " ist nicht existent!")
        GoTo end2_NewReklamationE
    Else
        'Öffnen der Fehlererfassung mit vorhandener Rekl-Nr.
        '20151002-RSP Filter direkt setzen
        'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_get_p_ReklNr ()l (), acFormEdit
        If GetStateVorgang(get_p_ReklNr()) = "Abgeschlossen" Then
'            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
            If IsLoaded("Fehlererfassung - Register") Then
                lb_reload = True
                Call unlockDataset(Forms("Fehlererfassung - Register").Controls("Mastertab#ID"))
            End If
            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormReadOnly
            If lb_reload Then Call Forms("Fehlererfassung - Register").ReLoadForm
        Else
            If proofAndLockDataset(get_p_ReklNr()) = "" Then
'                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit
                If IsLoaded("Fehlererfassung - Register") Then
                    lb_reload = True
                    Call unlockDataset(Forms("Fehlererfassung - Register").Controls("Mastertab#ID"))
                End If
                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
                If lb_reload Then Call Forms("Fehlererfassung - Register").ReLoadForm
                ReklBearbeiten = True
            Else
'                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
                If IsLoaded("Fehlererfassung - Register") Then
                    lb_reload = True
                    Call unlockDataset(Forms("Fehlererfassung - Register").Controls("Mastertab#ID"))
                End If
                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormReadOnly
                If lb_reload Then Call Forms("Fehlererfassung - Register").ReLoadForm
                ReklBearbeiten = False
            End If
        End If
    End If
End If

'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
If (ReklBearbeiten) Then
    InsertNewReklMemoDS (get_p_ReklNr())
    InsertNewEinkaufDS (get_p_ReklNr())
End If

If IsLoaded("Fehlererfassung - Register") Then
    Set lf_FE_Form = Forms("Fehlererfassung - Register")
    With Forms("Fehlererfassung - Register")
        .Requery
        .Refresh
        .AllowAdditions = False
        'DoCmd.Maximize
        'Sichern des aktuellen Vorgang-Status
        If IsNull(lf_FE_Form.Controls("Mastertab#Kategorie-Status")) Then
            p_VorgangStatus = ""
        Else
            p_VorgangStatus = lf_FE_Form.Controls("Mastertab#Kategorie-Status")
        End If
        'Forms("Fehlererfassung - Register").Controls("Kunden-ID").SetFocus  '20121121-RSP Änderung
    End With


    If Not IsNull(rs) Then Set rs = Nothing

    If lf_FE_Form.Controls("Mastertab#Datum#1") = #1/1/9999# Then
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Kurzzeichen] = '" & get_pKuerzel() & "';")
        ldb.Execute ("Update Mastertab set [Datum] = '" & Date & "', [Zuständiger Service-MA] = '" & rs("Name").Value & "', [Kategorie-Status]='Bearbeitung: Service'  where ID = " & p_ReklNr & ";")
        'lf_FE_Form.Controls ("")
        Test = LoadControlsFormAfterOpen(lf_FE_Form, lf_FE_Form.getRecordSource(), "[ID]=" & p_ReklNr)
        Test = LoadControlsPageAfterOpen(lf_FE_Form, lf_FE_Form.getRecordSource(), lf_FE_Form.Controls("RegisterStr246").Pages(lf_FE_Form.getActivePage()).Controls, "[ID]=" & p_ReklNr)
'            Forms("Fehlererfassung - Register").frmFRControlEventAfterUpdate ("RegisterStr246_Change")
        Set rs = Nothing
    End If
End If

If fReason <> "New" Then
    Forms("Fehlererfassung - Register").SetFocus
End If

end_NewReklamationE:
end2_NewReklamationE:


End Sub
Public Sub OpenReklamationE()

Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean
Dim ReklBearbeiten As Boolean

'Set db = currentdb ()
NewRekl = False
'MsgBox "Reklamation: " & ReklNr
'gibt es die eingegeben Reklamationsnummer?
Dim strKrit As String
strKrit = "ID = " & get_p_ReklNr()
If DCount("ID", "Mastertab", strKrit) = 0 Then
    MsgBox Replace(getTC_FCFromCode("TC0440"), "<xxx>", get_p_ReklNr())
    'MsgBox ("Reklamation mit der Nummer: " & get_get_p_ReklNr ()l () & " ist nicht existent!")
    GoTo end_OpenReklamationE
Else
    'Öffnen der Fehlererfassung mit vorhandener Rekl-Nr.
    'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_get_p_ReklNr ()l (), acFormEdit '20151002-RSP Filter direkt setzen
    If GetStateVorgang(get_p_ReklNr()) = "Abgeschlossen" Then
        DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly, acHidden
        'Forms("Fehlererfassung - Register").RecordsetType = 2 'Snapshot
    Else
        If proofAndLockDataset(get_p_ReklNr()) = "" Then
            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit, acHidden
            ReklBearbeiten = True
        Else
            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly, acHidden
            ReklBearbeiten = False
            'Forms("Fehlererfassung - Register").RecordsetType = 2 'Snapshot
        End If
    End If
End If
'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
If ReklBearbeiten Then
    InsertNewReklMemoDS (get_p_ReklNr())
End If
With Forms("Fehlererfassung - Register")
    .AllowAdditions = False
    'DoCmd.Maximize

    'Sichern des aktuellen Vorgang-Status
    If IsNull(Forms("Fehlererfassung - Register").Controls("Kategorie-Status")) Then
        p_VorgangStatus = ""
    Else
        p_VorgangStatus = Forms("Fehlererfassung - Register").Controls("Kategorie-Status")
    End If
    Forms("Fehlererfassung - Register").Controls("Kunden-ID").SetFocus  '20121121-RSP Änderung
End With

end_OpenReklamationE:
End Sub
Public Sub NewReklamationI()

Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean
Dim ReklBearbeiten As Boolean

'Set db = currentdb ()

If get_pKuerzel() = "" Then
    If setKuerzel() = "" Then
        GoTo end_NewReklamationI
    End If
End If
'Öffnen der Fehlererfassung mit neuer Rekl-Nr
DoCmd.OpenForm "FehlererfassungIntern", acNormal, , , acFormAdd
DoCmd.GoToRecord , , acNewRec
'Das aktuelle Datum wird als Auftragsdatum eingetragen
' -> Notwendig damit die Reklamtionsnummer sofort generiert wird und für weiter Zwecke genutzt werden kann
Forms("FehlererfassungIntern").Controls("Auftragsdatum") = format(Date, "dd.mm.yy")
Forms("FehlererfassungIntern").Controls("Status") = "Bearbeitung: QS-Fertigung"
p_ReklNr = Forms("FehlererfassungIntern").Controls("ID")
Forms("FehlererfassungIntern").Controls("Erfasser_Kurzzeichen") = getKundenIDfromKuerzel(get_pKuerzel())
If Forms("FehlererfassungIntern").Controls("Erfasser_Kurzzeichen") = 0 Then
    MsgBox Replace(getTC_FCFromCode("TC0441"), "<xxx>", get_pKuerzel())
    'MsgBox ("Das Kürzel '" & get_get_get_pKuerzel () () () & "' ist nicht bekannt, bitte den Erfasser eintragen!")
    'get_get_get_pKuerzel () () () = ""
End If
'Anlegen des Datensatz zum Sperren der Bearbeitung durch einen Zweiten Benutzer
lockDataset (get_p_ReklNr())
ReklBearbeiten = True
'
'InsertNewReklMemoDS (get_get_p_ReklNr ()l ())
'Beim Öffnen so setzen das der generierte Kommentar nicht kopiert wird. Kopie macht erst Sinn wenn Änderungen
'am Datensatz bzw. direkt im Kommentarfeld vollzogen wurden. In dem Fall wird die Variable wie folgt geändert:
'Bei einer Änderung im Steuerelement wird die Variable auf True gesetzt (
p_KommentarKopieren = False

NewRekl = True
'Ordner für die Dokumentenablage anlegen - Ordnername = ReklNr
With Forms("FehlererfassungIntern")
'    If Dir("F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\" & Format(.Auftragsdatum, "yyyy") & "\" & .ID & "\", vbDirectory) = "" Then
'        MkDir ("F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\" & Format(.Auftragsdatum, "yyyy") & "\" & .ID & "\")
'    End If
    If Dir(getTC_FCFromCode("DP0004") & format(.Auftragsdatum, "yyyy") & "\" & .ID & "\", vbDirectory) = "" Then
        MkDir (getTC_FCFromCode("DP0004") & format(.Auftragsdatum, "yyyy") & "\" & .ID & "\")
    End If
End With

'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
If ReklBearbeiten Then
    InsertNewReklMemoDS (get_p_ReklNr())
End If
With Forms("FehlererfassungIntern")
    'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
    .Filter = "[Mastertab].[ID] = " & get_p_ReklNr()
    .FilterOn = True
    .Requery
    .Refresh
    .AllowAdditions = False
    'Set Interner Fehler
    .intern = True
    '05.09.2013 -RSP: Ende
    'DoCmd.Maximize

    'Sichern des aktuellen Vorgang-Status
    If IsNull(Forms("FehlererfassungIntern").Controls("Kategorie-Status")) Then
        p_VorgangStatus = ""
    Else
        p_VorgangStatus = Forms("FehlererfassungIntern").Controls("Kategorie-Status")
    End If
    Forms("FehlererfassungIntern").Controls("KundenName").SetFocus  '20121121-RSP Änderung
End With


end_NewReklamationI:

End Sub

Public Sub OpenReklamationI()

Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean
Dim ReklBearbeiten As Boolean


If get_pKuerzel() = "" Then
    If setKuerzel() = "" Then
        GoTo end_OpenReklamationI
    End If
End If
'Set db = currentdb ()
NewRekl = False
'MsgBox "Reklamation: " & ReklNr
'gibt es die eingegeben Reklamationsnummer?
Dim strKrit As String
strKrit = "ID = " & get_p_ReklNr()
If DCount("ID", "Mastertab", strKrit) = 0 Then
    MsgBox Replace(getTC_FCFromCode("TC0440"), "<xxx>", get_p_ReklNr())
'    MsgBox ("Reklamation mit der Nummer: " & get_get_p_ReklNr ()l () & " ist nicht existent!")
    GoTo end_OpenReklamationI
Else
    'Öffnen der Fehlererfassung mit vorhandener Rekl-Nr.
    'DoCmd.OpenForm "FehlererfassungIntern", acNormal, , "[Mastertab].[ID] = " & get_get_p_ReklNr ()l (), acFormEdit        '20151002-RSP Filter direkt setzen
'    If GetStateVorgang(get_get_p_ReklNr ()l ()) = "Abgeschlossen" Then
'        Forms("FehlererfassungIntern").RecordsetType = 2 'Snapshot
'    Else
        If proofAndLockDataset(get_p_ReklNr()) = "" Then
'            DoCmd.OpenForm "FehlererfassungIntern", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit           '20151002-RSP Filter direkt setzen
'            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit, acHidden    '20151002-RSP Filter direkt setzen     '20190407-ROS
            ReklBearbeiten = True
        Else
'            DoCmd.OpenForm "FehlererfassungIntern", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
'            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit, , acHidden          '20151002-RSP Filter direkt setzen     '20190407-ROS
            ReklBearbeiten = False
            'Forms("FehlererfassungIntern").RecordsetType = 2 'Snapshot
        End If
'    End If
End If
'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
If ReklBearbeiten Then
    InsertNewReklMemoDS (get_p_ReklNr())
End If

DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit, acWindowNormal
'set_FehlererfassungEnviroment       '20190407-ROS
'With Forms("FehlererfassungIntern")
With Forms("Fehlererfassung - Register")        '20190407-ROS
    'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
'    .Filter = "[Mastertab].[ID] = " & get_get_p_ReklNr ()l ()
'    .FilterOn = True
    .Requery
    .Refresh
    .AllowAdditions = False
    '05.09.2013 -RSP: Ende
    'DoCmd.Maximize

    'Sichern des aktuellen Vorgang-Status
'    If IsNull(Forms("FehlererfassungIntern").Controls("Kategorie-Status")) Then
    If IsNull(Forms("Fehlererfassung - Register").Controls("Mastertab#Kategorie-Status")) Then        '20190407-ROS
        p_VorgangStatus = ""
    Else
'        p_VorgangStatus = Forms("FehlererfassungIntern").Controls("Kategorie-Status")
        p_VorgangStatus = Forms("Fehlererfassung - Register").Controls("Mastertab#Kategorie-Status")        '20190407-ROS
    End If
'    Forms("FehlererfassungIntern").Controls("KundenName").SetFocus  '20121121-RSP Änderung
'    Forms("Fehlererfassung - Register").Controls("KundenName").SetFocus  '20121121-RSP Änderung        '20190407-ROS
End With

end_OpenReklamationI:
End Sub

Function getDateAddTage(D As Date, ByVal t As Integer) As Date

    Dim X As Integer 'offset Datum
    Dim i As Integer

    getDateAddTage = D
    X = 0
    i = 1
    While (i <= t + X)
    'For i = 1 To t + x
        Select Case Weekday(D + i, vbMonday)
            Case 1, 2, 3, 4, 5 'Mo-Fr
                getDateAddTage = getDateAddTage + 1
            Case 6, 7 'Sa-So
                X = X + 1
                getDateAddTage = getDateAddTage + 1
        End Select
        i = i + 1
    Wend
    'Next i

End Function

Function getDelayTwoDates(DFlex As Date, DStat As Date) As Long
'Beschreibung:
'Paramter:  DFlex - Heute
'           DStat = Eingestelltes Datum (Referenz)
'Funktion berechnet die Anzahl Arbeitstage zwischen dem Startdatum DFlex und dem Endedatum DStat (Arbeitstage = DStat - DFlex)

    Dim anzWeek As Integer
    Dim varDatum As Variant
    Dim AnzTage As Integer

    If DFlex >= DStat Then
        'Anzahl Wochen
        anzWeek = DateDiff("w", DStat, DFlex)
        AnzTage = (anzWeek) * 5
        varDatum = DateAdd("ww", anzWeek, DStat)

        Do While varDatum < DFlex
            If Weekday(varDatum) <> vbSunday _
            And Weekday(varDatum) <> vbSaturday Then
                AnzTage = AnzTage + 1
            End If
            'Nächstes zuprüfendes Datum
            varDatum = DateAdd("d", 1, varDatum)
        Loop
        'Forms("BerechnenWerktageDiffDatum").Controls("WT") = AnzTage
        getDelayTwoDates = AnzTage
    Else
        getDelayTwoDates = DFlex - DStat
    End If

End Function

Function CalcArbeitszeitDiffDate(fDateStart As Date, fDateEnd As Date) As Long   '20170220-ROS
'Beschreibung:
'Paramter: fDateStart   - Startdatum und Startzeit
'Paramter: fDateEnd     - Endedatum und Endezeit
'Funktion berechnet die Arbeitszeit zwischen den beiden Parametern,
' Dabei werden die Wochenenden nur berücksichtigt wenn entweder Start oder Endedatum auf einen Samstag fallen.
' Zwischen den Datums werden die Wochenenden nicht berücksicht.
' Alle Wochentage werden mit 9 Stunden inkl. Pausen angenommen: Arbeitszeit 07:30 - 16:30
' Der Samstag wird, wenn einer der Beiden Paratemer ein Sa ist mit 5 Stunden berücksichtigt: 07:30 - 12:30
Dim liDiffT1Min As Long
Dim liDiffT2Min As Long
Dim liAnzTage As Long
Dim liZeitDiffGesMin As Long

CalcArbeitszeitDiffDate = 0
'If Format(fDateStart, "dd:mm:yyyy hh:mm") < Format(fDateEnd, "dd:mm:yyyy hh:mm") Then ' EndeDatum/Zeit muss immer größer sein
    If Weekday(fDateStart) = vbSaturday Then
        liDiffT1Min = 750 - (Hour(format(fDateStart, "hh:mm")) * 60 + Minute(format(fDateStart, "hh:mm"))) 'Anzahl MInuten bis 12:30 Uhr Ende - Anzahl Minuten Start
    Else
        'Wenn Starttime (fDateStart) >= 16:30 Uhr dann DIff = 0 setzen
        If format(fDateStart, "hh:mm") >= "16:30" Then
            liDiffT1Min = 0
        Else
            liDiffT1Min = 990 - (Hour(format(fDateStart, "hh:mm")) * 60 + Minute(format(fDateStart, "hh:mm"))) 'Anzahl MInuten bis 16:30 Uhr Ende - Anzahl Minuten Star
        End If
    End If
    'Wenn Endtime (fDateEnd) <= 07:30 Uhr dann DIff = 0 setzen, da die Zeit sich vor dem Arbeitsbeginn befindet
    If format(fDateEnd, "hh:mm") <= "07:30" Then
        liDiffT2Min = 0
    Else
        liDiffT2Min = (Hour(format(fDateEnd, "hh:mm")) * 60 + Minute(format(fDateEnd, "hh:mm"))) - 450 'Anzahl Minuten Ende - Anzahl MInuten bis 07:30 Uhr Start
    End If
    'Abfrage ob die EndZeit größer ist als die Startzeit - In Abhängigkeit ob das Ende einen Tag später ist, handelt sich Diffrenz um mehr als 9 Stunden
    If format(fDateEnd, "hh:mm:ss") > format(fDateStart, "hh:mm:ss") Then
        'Vorgang ist über einen Tag laut Datum
        If format(fDateEnd, "dd:mm:yyyy") > format(fDateStart, "dd:mm:yyyy") Then
            liAnzTage = getDelayTwoDates(fDateEnd, fDateStart)
            'Wenn das Endedatum kein Samstag ist muss ein Tag abgezogen werden, da die Minuten schon für den letzten Tag berechnet sind.
            'Bei Samstag muss der Tag berücksichtigt werden da in der Funktion getDelayTwoDates der Samstag abgezogen wird
            If Weekday(fDateEnd) <> vbSaturday Then liAnzTage = liAnzTage - 1
            If liAnzTage = 0 Then
                liZeitDiffGesMin = liDiffT1Min + liDiffT2Min
            Else
                If Weekday(fDateStart) <> vbSaturday Then liAnzTage = liAnzTage - 1
                liZeitDiffGesMin = ((liAnzTage) * 540) + (liDiffT1Min + liDiffT2Min)
                'liZeitDiffGesMin = ((liAnzTage - 1) * 540) + (liDiffT1Min + liDiffT2Min)
            End If
        Else
            'Gleicher Tag
            If format(fDateEnd, "dd:mm:yyyy") = format(fDateStart, "dd:mm:yyyy") Then
                If Weekday(fDateStart) = vbSaturday Then
                    liZeitDiffGesMin = (liDiffT1Min + liDiffT2Min) - 300
                Else
                    liZeitDiffGesMin = (liDiffT1Min + liDiffT2Min) - 540
                End If
'            Else
'                MsgBox "EndeDatum kann nicht kleiner sein als StartDatum!", vbCritical
            End If
        End If
    Else
        'If Format(fDateEnd, "dd:mm:yyyy") > Format(fDateStart, "dd:mm:yyyy") Then
        If fDateEnd > fDateStart Then
            liAnzTage = getDelayTwoDates(fDateEnd, fDateStart)
            If Weekday(fDateStart) = vbSaturday Then
                liAnzTage = liAnzTage + 1
            End If
            liZeitDiffGesMin = ((liAnzTage - 1) * 540) + (liDiffT1Min + liDiffT2Min)
        End If

    End If
    CalcArbeitszeitDiffDate = liZeitDiffGesMin
'End If

End Function


Function getHourMinuteByAnzMinute(fiAnzMinute, ByRef fiDays, ByRef fiHour, ByRef fiMinute) As Double    '20170220-ROS

    Dim liDays As Double

    fiHour = (fiAnzMinute Mod 540) \ 60
    fiMinute = fiAnzMinute Mod 60
    fiDays = fiAnzMinute \ 540
    liDays = fiAnzMinute \ 540
    liDays = liDays + ((fiAnzMinute Mod 540) \ 54) / 10

    getHourMinuteByAnzMinute = liDays





End Function


Sub sendMailLieferant(ByVal absender As String, ByVal empfaenger As String, ByVal Betreff As String, ByVal body As String, ByVal Anhang As String, ByVal Anhang2 As String, ByVal MailOeffnen As Boolean, Optional ByVal CC As String = "", Optional ByVal BCC As String = "")

    Dim iMsg As Object
    Dim iConf As Object
    Dim Flds

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-intern.link.local"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "smtpsend@link.local"   '"LQ Admin"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "S3nDMa!L11"   '"123456"
        .Update
    End With

    '##### Ermitteln Empfänger sowie Absender Account
    With iMsg
        Set .Configuration = iConf
        .To = empfaenger
        .CC = CC
        If (getCurrLanguage = "German") Then
            .BCC = absender & "; service@de.lq-group.com"
        Else
            .BCC = absender
        End If
        .From = absender
        .subject = Betreff
        '.TextBody = strbody
        .AddAttachment Anhang
        'If Anhnag2 <> "" Then
        '    .AddAttachment "C:\Users\ralf sprotte\Documents\FH.pdf"
        'End If
        .TextBody = body
'        If MailOeffnen Then
'            .Display
'        Else
            .send
'        End If
        '.Send
    End With

    Set iConf = Nothing
    Set iMsg = Nothing

End Sub
Sub sendMailLieferantAnhang(ByVal absender As String, ByVal empfaenger As String, ByVal Betreff As String, ByVal body As String, Anhang() As String, ByVal MailOeffnen As Boolean, Optional ByVal CC As String = "", Optional ByVal BCC As String = "")
On Error GoTo err_sendMailLieferantAnhang
    Dim iMsg As Object
    Dim iConf As Object
    Dim Flds
    Dim i As Integer

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-intern.link.local"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "smtpsend@link.local"   '"LQ Admin"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "S3nDMa!L11"   '"123456"
        .Update
    End With

    '##### Ermitteln Empfänger sowie Absender Account
    With iMsg
        Set .Configuration = iConf
        .To = empfaenger
        .CC = CC
        If (getCurrLanguage = "German") Then
            .BCC = absender & "; service@de.lq-group.com"
        Else
            .BCC = absender
        End If
        .From = absender
        .subject = Betreff
        '.TextBody = strbody
        For i = 0 To UBound(Anhang) - 1
            .AddAttachment Anhang(i)
        Next i

'        .AddAttachment Anhang
'        If Anhang2 <> "" Then
'            .AddAttachment Anhang2
'        End If
        .TextBody = body
'        .HTMLBody = Body
'        If MailOeffnen Then
'            .Display
'        Else
            DoCmd.Hourglass True
            .send
            DoCmd.Hourglass False
'        End If
        '.Send
    End With
end_sendMailLieferantAnhang:
    Set iConf = Nothing
    Set iMsg = Nothing
    Exit Sub
err_sendMailLieferantAnhang:
    MsgBox Err.Description
    GoTo end_sendMailLieferantAnhang

End Sub
Sub sendMailQSV(ByVal absender As String, ByVal empfaenger As String, ByVal Betreff As String, ByVal body As String, byValbearbeitunsdatum As Date, Optional ByVal CC As String = "", Optional ByVal BCC As String = "")

    Dim iMsg As Object
    Dim iConf As Object
    Dim Flds


    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-intern.link.local"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "smtpsend@link.local"   '"LQ Admin"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "S3nDMa!L11"   '"123456"
        .Update
    End With

    '##### Ermitteln Empfänger sowie Absender Account
    With iMsg
        Set .Configuration = iConf
        .To = empfaenger
        .CC = CC
        .BCC = ""
        .From = absender
        .subject = Betreff
        '.TextBody = strbody
'        .TextBody = Body
        .HTMLBody = body  '20161215-ROS
        .send
    End With

    Set iConf = Nothing
    Set iMsg = Nothing

End Sub

Function CheckNextMahnstufeLieferant(ByVal frs As DAO.Recordset, ByVal fdb As DAO.Database, fMahnstufe As String) As Boolean
On Error GoTo err_CheckNextMahnstufeLieferant

        Dim rs, rs2 As DAO.Recordset
        Dim strBody As String
        Dim sVon As String
        Dim sTo As String
        Dim scc As String
        Dim sbcc As String
        Dim sBetreff As String
        Dim sBody As String
        Dim sAnhang As String
        Dim sPfadDatei As String
        Dim BearbeitungDatumLief As Date
        Dim sSBMail As String
        Dim sSBTel As String


        CheckNextMahnstufeLieferant = True

        '##### Ermitteln Empfänger sowie Absender Account
        Set rs = fdb.OpenRecordset("Select * from RSP_AnsprechpartnerLiefUndEinkauf where AnsprechpartnerID = " & frs("LiefAnsprechp-ID") & ";")
        If rs.RecordCount > 0 Then
            sTo = rs("Email")
        Else
            MsgBox getTC_FCFromCode("TC0296") & " - " & getTC_FCFromCode("TC0450") & ".: " & frs("ID") '("Kein Ansprechpartner eingetragen")
        End If
        If fMahnstufe = "2.Mahnung" Then  '20160706-ROS
            If Not IsNull(rs("EmailName")) Or rs("EmailName") = "" Then
                scc = rs("EmailName") 'Email Sachbearbeiter Einkauf
            Else
                scc = getMahnstufeText("EinkaufAccount2teMahnung") '(Zentral E-Mail Einkauf)
            End If
        Else
            scc = ""
        End If
        sbcc = ""
        Set rs = fdb.OpenRecordset("Select * from Sachbearbeiter where Name='" & frs("Sachbearbeiter") & "';")
        If rs.RecordCount > 0 Then
            sVon = rs("EmailName")
            sSBMail = rs("EmailName")
            sSBTel = rs("TelefonBeruflich")
        Else
            MsgBox getTC_FCFromCode("TC0295") 'MsgBox ("Kein Sachbearbeiter eingetragen")
        End If
        Set rs = Nothing

        sBetreff = fMahnstufe & " " & getTC_FCFromCode("TC0418") & " " & CStr(frs("ID")) ' & rmaNr
        If IsNull(frs("Meldungsnummer Kunde")) Or frs("Meldungsnummer Kunde") = "" Then
            sPfadDatei = setReklVerzeichnis(frs("ID"), format(frs("Auftragsdatum"), "yyyy"))
        Else
            sPfadDatei = setReklVerzeichnis(frs("ID"), format(frs("Auftragsdatum"), "yyyy"), frs("Meldungsnummer Kunde"))
        End If

        sAnhang = sPfadDatei & CStr(frs("ID")) & "_Reklamationsbericht_Lieferant.pdf"

'20160720-ROS
        sBody = getTC_FCFromCode("TC0442") & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0443") & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0444") & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0445") & " " & frs("Rückmeldung_Lieferant_Verlängertl") & " " & getTC_FCFromCode("TC0462") & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0446") & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0447") & vbNewLine & _
                getTC_FCFromCode("TC0448") & vbNewLine & vbNewLine & _
                sSBMail & " " & getTC_FCFromCode("TC0449") & " " & sSBTel & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0455") & vbNewLine & _
                getTC_FCFromCode("TC0454")

'                "LQ Mechatronik-Systeme GmbH" & vbNewLine & _
'                "Carl-Benz-Straße 6 _D-74354 Besigheim"
'        sTo = "ralf.sprotte@de.lq-group.com"
'        sVon = "service@de.lq-group.com"
        sendMailLieferant sVon, sTo, sBetreff, sBody, sAnhang, "", False, scc, sbcc
        Set rs = Nothing

        Exit Function
err_CheckNextMahnstufeLieferant:

    If Err = 3021 Then
        CheckNextMahnstufeLieferant = False
        Exit Function
    End If

End Function
Function CheckNextMahnstufeQSV(ByVal frs As DAO.Recordset, ByVal fdb As DAO.Database, ByVal fMahnstufe As String) As Boolean
On Error GoTo err_CheckNextMahnstufeQSV

        Dim ldb As DAO.Database
        Dim rs As DAO.Recordset
        Dim rsID As DAO.Recordset
        Dim strBody As String
        Dim sVon As String
        Dim sTo As String
        Dim scc As String
        Dim sbcc As String
        Dim sBetreff As String
        Dim sBody As String
        Dim sAnhang As String
        Dim sPfadDatei As String
        Dim dBearbeitungDatum As Date
        Dim bMailEinheit As Boolean
        Dim sSave As String
        Dim sAnschreiben As String
        Dim sAnrede As String
        Dim sBodyHTML As String

        CheckNextMahnstufeQSV = True

         Set ldb = openAccDBConnection("openDB")
        '##### Ermitteln Empfänger sowie Absender Account
        bMailEinheit = False
        If Not IsNull(frs("LQBereichID")) And Not frs("LQBereichID") = -1 Then
            Set rs = ldb.OpenRecordset("Select * from LQBereich where [ID]=" & frs("LQBereichID") & ";", dbOpenSnapshot)
            If rs("MailTOLQBereich") = "" Or IsNull(rs("MailTOLQBereich")) Then
                bMailEinheit = True
            Else
                sTo = rs("MailTOLQBereich")
                scc = rs("MailCCLQBereich")
                sbcc = rs("MailEskalationLQBereich")
            End If
        Else
            bMailEinheit = True
        End If

        If bMailEinheit Then
            Set rs = ldb.OpenRecordset("Select * from LQEinheit where [ID]=" & frs("LQEinheitID") & ";", dbOpenSnapshot)
            If rs("MailTOLQEinheit") = "" Or IsNull(rs("MailTOLQEinheit")) Then
                MsgBox getTC_FCFromCode("TC0251") '("Es kann keine Mail an die LQ-Einheit gesendet werden, weil keine Mail hinterlegt ist!")
                GoTo exit_CheckNextMahnstufeQSV
            Else
                sTo = rs("MailTOLQEinheit")
                scc = rs("MailCCLQEinheit")
                sbcc = rs("MailEskalationLQEinheit")
            End If

        End If
        Set rs = Nothing
        If fMahnstufe = "2. Mahnung zur Reklamation" Then
            sSave = sTo
            sTo = sbcc 'Email an Eskalation
            scc = sSave
        Else
            sbcc = ""
        End If

        If (getCurrLanguage = "German") Then
            sVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
        Else
            sVon = ""
        End If


        'sBetreff = fMahnstufe & "" & CStr(frs("ID"))
        sBetreff = Replace(cBetreff, "!Textwert!", fMahnstufe)
        sBetreff = Replace(sBetreff, "!ReklNr!", frs("ID")) & getTC_FCFromCode("TC0368")
        sAnschreiben = getTC_FCFromCode("TC0370") & " " & frs("ID") & " " & getTC_FCFromCode("TC0412") & " " & frs("QSBearbeitungBis_Verl") & " " & getTC_FCFromCode("TC0409") & "."

        If Not IsNull(frs("Sachbearbeiter")) Then
            sAnrede = getTC_FCFromCode("TC0424") & " " & frs("Sachbearbeiter") & ","
        Else
            sAnrede = getTC_FCFromCode("TC0424") & ","
        End If

        sBody = sAnrede & vbNewLine & vbNewLine & sAnschreiben & vbNewLine & vbNewLine & _
                cTrennStrich & vbNewLine & getTC_FCFromCode("TC0371") & vbNewLine & vbNewLine & frs("Welcher Fehler tritt auf") & vbNewLine & vbNewLine & cTrennStrich & vbNewLine & vbNewLine & _
                getTC_FCFromCode("TC0372") & vbNewLine & cTrennStrich & vbNewLine & getTC_FCFromCode("TC0373") & _
                vbNewLine & vbNewLine & vbNewLine & vbNewLine & cTrennStrich & vbNewLine & getTC_FCFromCode("TC0374") & vbNewLine & vbNewLine & vbNewLine & vbNewLine
        '20161215-ROS
        sBodyHTML = "<html><body>" & sAnrede & "<br>" & "<br>" & sAnschreiben & "<br>" & "<br>" & _
                cTrennStrich & "<br>" & getTC_FCFromCode("TC0371") & "<br>" & "<br>" & frs("Welcher Fehler tritt auf") & "<br>" & "<br>" & cTrennStrich & "<br>" & "<br>" & _
                getTC_FCFromCode("TC0372") & "<br>" & cTrennStrich & "<br>" & getTC_FCFromCode("TC0373") & _
                "<br>" & "<br>" & "<br>" & "<br>" & cTrennStrich & "<br>" & getTC_FCFromCode("TC0374") & "<br>" & "<br>" & "<br>" & "<br>" & _
                "<a href=""" & sPfadDatei & """>" & getTC_FCFromCode("TC0482") & "</a></body></html>"
        dBearbeitungDatum = frs("QSBearbeitungBis_Verl")

        sendMailQSV sVon, sTo, sBetreff, sBodyHTML, dBearbeitungDatum, scc, sbcc

        Set rs = Nothing

exit_CheckNextMahnstufeQSV:
        Set rs = Nothing
        Exit Function

err_CheckNextMahnstufeQSV:
        MsgBox getTC_FCFromCode("TC0477")
        'MsgBox ("Fehler in CheckNextMahnstufeQSV!")
        CheckNextMahnstufeQSV = False
        GoTo exit_CheckNextMahnstufeQSV

End Function


Sub CheckAndSendMahnungenLieferant()

'SELECT Mastertab.ID, Mastertab.[MahnungLief-ID], LieferantMahnwesen.[Mahnustufe-ID], Mahnstufen.Mahnstufe, LieferantMahnwesen.Datum, Mastertab.Rückmeldung_Lieferant_Verlängertl, Mastertab.Rückmeldung_Lieferant_Ist, Mastertab.Rückmeldung_Lieferant_Verlängertl, Date()-[Rückmeldung_lieferant_Verlängertl] AS Delay
'FROM (Mastertab INNER JOIN LieferantMahnwesen ON Mastertab.[MahnungLief-ID] = LieferantMahnwesen.[Mahnung-ID]) INNER JOIN Mahnstufen ON LieferantMahnwesen.[Mahnustufe-ID] = Mahnstufen.[MS-ID]
'WHERE (((Mastertab.Rückmeldung_Lieferant_Verlängertl) Is Not Null) AND ((Mastertab.Rückmeldung_Lieferant_Ist) Is Null));

Dim ldb As DAO.Database
Dim rs, rs2, rsFilter, rsKonf, rsMS, rsMS2 As DAO.Recordset
Dim sSQLState As String

Dim iMsg As Object
Dim iConf As Object
Dim strBody, sVon, sTo, scc, sbcc, sBetreff, sBody, sAnhang, sPfadDatei As String

sSQLState = "SELECT Mastertab.ID, Mastertab.[MahnungLief-ID], LieferantUndQSVMahnwesen.[Mahnustufe-ID], Mahnstufen.Mahnstufe, " & _
            "LieferantUndQSVMahnwesen.Datum, Mastertab.Rückmeldung_Lieferant_Verlängertl, Mastertab.Rückmeldung_Lieferant_Ist, " & _
            "Mastertab.Rückmeldung_Lieferant_Verlängertl, getDelayTwoDates(Date(),[Rückmeldung_lieferant_Verlängertl]) AS Delay, " & _
            "Mastertab.[LiefAnsprechp-ID], Mastertab.Sachbearbeiter, Mastertab.[Meldungsnummer Kunde], Mastertab.Auftragsdatum, Mastertab.[Kategorie-Status] " & _
            "FROM (Mastertab INNER JOIN LieferantUndQSVMahnwesen ON Mastertab.[MahnungLief-ID] = LieferantUndQSVMahnwesen.[Mahnung-ID]) " & _
            "INNER JOIN Mahnstufen ON LieferantUndQSVMahnwesen.[Mahnustufe-ID] = Mahnstufen.[MS-ID] " & _
            "WHERE (((Mastertab.[Kategorie-Status])<>'Abgeschlossen') AND ((Mahnstufen.Mahnstufe)<>'XD-Report mit Befund') AND " & _
            "((Mastertab.Rückmeldung_Lieferant_Verlängertl) Is Not Null) AND ((Mastertab.Rückmeldung_Lieferant_Ist) Is Null) AND " & _
            "((LieferantUndQSVMahnwesen.Typ)='Lieferant') AND ((getDelayTwoDates(Date(),[Rückmeldung_lieferant_Verlängertl]))>="


Set ldb = openAccDBConnection("openDB")
Set rsMS = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = 'Erinnerung';", dbOpenSnapshot)
Set rsMS2 = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = 'XD-Report versendet';", dbOpenSnapshot)
'Prüft alle Datensätze deren Fälligkeit Erinnerung erreicht oder älter sind
Set rs = ldb.OpenRecordset(sSQLState & rsMS("Faelligkeit") & "));") '!!!!!!!  Korrekter Filter für Erinnerung Wert 1
If rs.RecordCount > 0 Then
    rs.MoveFirst
    'Über alle gefundenen Datensätze
    While Not rs.EOF
'        If rs("ID") = 28442 Or rs("ID") = 28503 Then
'            Debug.Print rs("ID")
'        End If
        'Wurde für diesen Vorgang schon eine Erinnerung gesendet? Abfrage Status = "XD-Report versendet" => muss noch eine Erinnerung gesendet werden.
        If rs("Mahnustufe-ID") = rsMS2("MS-ID") Then
            If CheckNextMahnstufeLieferant(rs, ldb, "BEARBEITUNGSERINNERUNG") Then
                rs.Edit
                rs("Mahnustufe-ID") = rsMS("MS-ID") 'Erinnerung
                rs("Datum") = Date
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
End If

Set rsMS = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = '1. Mahnung';", dbOpenSnapshot)
Set rsMS2 = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = 'Erinnerung';", dbOpenSnapshot)
'Prüft alle Datensätze deren Fälligkeit 1.Mahnstufe erreicht haben oder älter sind
Set rs = ldb.OpenRecordset(sSQLState & rsMS("Faelligkeit") & "));") '!!!!!!!  Korrekter Filter für Erinnerung Wert 3
If rs.RecordCount > 0 Then
    rs.MoveFirst
    'Über alle gefundenen Datensätze
    While Not rs.EOF
'        If rs("ID") = 28442 Or rs("ID") = 28503 Then
'            Debug.Print rs("ID")
'        End If
        If rs("Mahnustufe-ID") = rsMS2("MS-ID") Then
            If CheckNextMahnstufeLieferant(rs, ldb, "1.Mahnung") Then
                rs.Edit
                rs("Mahnustufe-ID") = rsMS("MS-ID") '1.Mahnung
                rs("Datum") = Date
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
End If


Set rsMS = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = '2. Mahnung';", dbOpenSnapshot)
Set rsMS2 = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = '1. Mahnung';", dbOpenSnapshot)
'Prüft alle Datensätze deren Fälligkeit die 2.Mahnstufe erreicht haben oder älter sind
Set rs = ldb.OpenRecordset(sSQLState & rsMS("Faelligkeit") & "));") '!!!!!!!  Korrekter Filter für Erinnerung Wert 5
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
'        If rs("ID") = 28442 Or rs("ID") = 28503 Then
'            Debug.Print rs("ID")
'        End If
        If rs("Mahnustufe-ID") = rsMS2("MS-ID") Then
            If CheckNextMahnstufeLieferant(rs, ldb, "2.Mahnung") Then
                rs.Edit
                rs("Mahnustufe-ID") = rsMS("MS-ID") '2.Mahnung
                rs("Kategorie-Status") = "Klärung: Lieferant/Einkauf"
                rs("Datum") = Date
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
End If


Set rsMS = Nothing
Set rsMS2 = Nothing
Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")


End Sub
Sub CheckAndSendMahnungenQSV()

'SELECT Mastertab.ID, Mastertab.[MahnungLief-ID], LieferantMahnwesen.[Mahnustufe-ID], Mahnstufen.Mahnstufe, LieferantMahnwesen.Datum, Mastertab.Rückmeldung_Lieferant_Verlängertl, Mastertab.Rückmeldung_Lieferant_Ist, Mastertab.Rückmeldung_Lieferant_Verlängertl, Date()-[Rückmeldung_lieferant_Verlängertl] AS Delay
'FROM (Mastertab INNER JOIN LieferantMahnwesen ON Mastertab.[MahnungLief-ID] = LieferantMahnwesen.[Mahnung-ID]) INNER JOIN Mahnstufen ON LieferantMahnwesen.[Mahnustufe-ID] = Mahnstufen.[MS-ID]
'WHERE (((Mastertab.Rückmeldung_Lieferant_Verlängertl) Is Not Null) AND ((Mastertab.Rückmeldung_Lieferant_Ist) Is Null));

Dim ldb As DAO.Database
Dim rs, rs2, rsFilter, rsKonf, rsMS, rsMS2 As DAO.Recordset
Dim sSQLState As String

Dim iMsg As Object
Dim iConf As Object
Dim strBody, sVon, sTo, scc, sbcc, sBetreff, sBody, sAnhang, sPfadDatei As String


sSQLState = "SELECT RSP_Q_MahnstufeMastertab.ID, RSP_Q_MahnstufeMastertab.LQEinheitID, RSP_Q_MahnstufeMastertab.LQBereichID, RSP_Q_MahnstufeMastertab.Sachbearbeiter, RSP_Q_MahnstufeMastertab.[Welcher Fehler tritt auf], " & _
            "RSP_Q_MahnstufeMastertab.QSBearbeitungBis_Ist, LieferantUndQSVMahnwesen.Datum, LieferantUndQSVMahnwesen.[Mahnustufe-ID], Mahnstufen.Mahnstufe, RSP_Q_MahnstufeMastertab.QSBearbeitungBis_Verl, " & _
            "getDelayTwoDates(Date(),[QSBearbeitungBis_Verl]) AS Delay, RSP_Q_MahnstufeMastertab.QSBearbeitungBis_Soll " & _
            "FROM (RSP_Q_MahnstufeMastertab INNER JOIN LieferantUndQSVMahnwesen ON RSP_Q_MahnstufeMastertab.ID = LieferantUndQSVMahnwesen.[R-ID]) " & _
            "INNER JOIN Mahnstufen ON LieferantUndQSVMahnwesen.[Mahnustufe-ID] = Mahnstufen.[MS-ID] " & _
            "WHERE (((RSP_Q_MahnstufeMastertab.[Kategorie-Status])<>'Abgeschlossen') AND ((Mahnstufen.Mahnstufe)<>'QS-Bearbeitung abgeschlossen') AND " & _
            "((RSP_Q_MahnstufeMastertab.QSBearbeitungBis_Verl) Is Not Null) AND ((LieferantUndQSVMahnwesen.Typ)='QSV') AND ((getDelayTwoDates(Date(),[QSBearbeitungBis_Verl]))>="


 Set ldb = openAccDBConnection("openDB")
Set rsMS = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = 'Erinnerung QSV';", dbOpenSnapshot)
Set rsMS2 = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = 'Versendet';", dbOpenSnapshot)
Set rs = ldb.OpenRecordset(sSQLState & rsMS("Faelligkeit") & "));") '!!!!!!!  Korrekter Filter für Erinnerung Wert 0
If rs.RecordCount > 0 Then
'    If rs("ID") = 28645 Or rs("ID") = 28553 Or rs("ID") = 28563 Then
'        Debug.Print rs("ID")
'    End If
    rs.MoveFirst
    While Not rs.EOF
        If rs("Mahnustufe-ID") = rsMS2("MS-ID") Then
            If CheckNextMahnstufeQSV(rs, ldb, rsMS("Textwert")) Then
                rs.Edit
                rs("Mahnustufe-ID") = rsMS("MS-ID") 'Erinnerung QSV
                rs("Datum") = Date
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
End If
Set rsMS = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = '1. Mahnung QSV';", dbOpenSnapshot)
Set rsMS2 = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = 'Erinnerung QSV';", dbOpenSnapshot)
Set rs = ldb.OpenRecordset(sSQLState & rsMS("Faelligkeit") & "));") '!!!!!!!  Korrekter Filter für 1. Mahnung Wert 2
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        If rs("Mahnustufe-ID") = rsMS2("MS-ID") Then
            If CheckNextMahnstufeQSV(rs, ldb, rsMS("Textwert")) Then
                rs.Edit
                rs("Mahnustufe-ID") = rsMS("MS-ID") '1.Mahnung
                rs("Datum") = Date
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
End If
Set rsMS = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = '2. Mahnung QSV';", dbOpenSnapshot)
Set rsMS2 = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe = '1. Mahnung QSV';", dbOpenSnapshot)
'Set rs = ldb.OpenRecordset(sSQLStateOhneDelay)
Set rs = ldb.OpenRecordset(sSQLState & rsMS("Faelligkeit") & "));") '!!!!!!!  Korrekter Filter für Erinnerung Wert 5
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        If rs("Mahnustufe-ID") = rsMS2("MS-ID") Then
            If CheckNextMahnstufeQSV(rs, ldb, rsMS("Textwert")) Then
                rs.Edit
                rs("Mahnustufe-ID") = rsMS("MS-ID") '2.Mahnung
                rs("Datum") = Date
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
End If

Set rsMS = Nothing
Set rsMS2 = Nothing
Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")


End Sub

Public Function getMahnstufeWerte(fMS As String) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe ='" & fMS & "';", dbOpenSnapshot)

    getMahnstufeWerte = CStr(rs("Faelligkeit"))

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

End Function

Public Function setMahnstufeWerte(fFaelligkeit As String, fMS As String) As Long

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

     Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Update Mahnstufen set [Faelligkeit] = " & CLng(fFaelligkeit) & " where [Mahnstufe]= '" & fMS & "';")

    setMahnstufeWerte = CLng(fFaelligkeit)

     ' Set ldb = openAccDBConnection("closeDB")

End Function

Public Function getMahnstufeText(fText As String) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Mahnstufen where Mahnstufe ='" & fText & "';", dbOpenSnapshot)

    If IsNull(rs("Textwert")) Then
        getMahnstufeText = ""
    Else
        getMahnstufeText = rs("Textwert")
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

End Function

Public Function setMahnstufeText(fSetText As String, fText As String) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Update Mahnstufen set [Textwert] = '" & fSetText & "' where [Mahnstufe]= '" & fText & "';")

    setMahnstufeText = fSetText

     ' Set ldb = openAccDBConnection("closeDB")

End Function


Public Sub NewPosAufwand(ByVal fReklID As Long, fKuerzel As String)
On Error GoTo err_NewPosAufwand

    Dim ldb As DAO.Database
    Dim rs, rsKuerzel As DAO.Recordset
    Dim sLastPos As String
    Dim lStdSatz As Integer  'Standardstundensatz
    Dim lStdSatzId As Integer   'Standartstundensatz ID

    lStdSatz = Nz(DLookup("Betrag", "tblAufwandStdSaetze", "[Standard] = True"))
    lStdSatzId = Nz(DLookup("StdSatzID", "tblAufwandStdSaetze", "[Standard] = True"))

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from AufwandVorgangPositionen where MID=" & fReklID & " ORDER BY VorgangID;")
    If rs.RecordCount = 0 Then
        'Anlegen Erfassungsposition 10
        rs.AddNew
        rs("MID") = fReklID
        rs("Pos") = "10"
        rs("Datum") = Date
        rs("Zeit") = 0.25
        rs("Vorgang") = 75 'ID entspricht 'Vorgang erfasst.'
        rs("StdSatz") = lStdSatz
        rs("StdSatzID") = lStdSatzId
        rs("DatumStdSatz") = Date
        rs("Aufwand") = 0.25 * lStdSatz
        Set rsKuerzel = ldb.OpenRecordset("Select Kurzzeichen from Bearbeiter where Name = '" & fKuerzel & "';")
        If rsKuerzel.RecordCount Then
            rs("Kürzel") = rsKuerzel("Kurzzeichen")
            rsKuerzel.Clone
            Set rsKuerzel = Nothing
        End If
        rs.Update
        'Den Aufwand der Reklamationsanlage in das Feld der Summer aller Aufwände eintragen
        ldb.Execute ("Update Mastertab set [Zeitbedarf] = 0.25 where [ID] = " & fReklID & ";")
    End If
    rs.MoveLast

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")
    Exit Sub
err_NewPosAufwand:
    Resume


End Sub
Public Sub NewPosAufwand2(ByVal fReklID As Long, fKuerzel As String) '20160205-ROS neue Angelegt
On Error GoTo err_NewPosAufwand2

    Dim ldb As DAO.Database
    Dim rs, rsKuerzel As DAO.Recordset
    Dim sLastPos As String
    Dim lStdSatz As Integer  'Standardstundensatz
    Dim lStdSatzId As Integer   'Standartstundensatz ID

    lStdSatz = Nz(DLookup("Betrag", "tblAufwandStdSaetze", "[Standard] = True"))
    lStdSatzId = Nz(DLookup("StdSatzID", "tblAufwandStdSaetze", "[Standard] = True"))

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from AufwandVorgangPositionen where MID=" & fReklID & " ORDER BY VorgangID;")
    If rs.RecordCount = 0 Then
        'Anlegen Erfassungsposition 10
        rs.AddNew
        rs("MID") = fReklID
        rs("Pos") = "10"
        rs("Datum") = Date
        rs("Zeit") = 0.25
        rs("Vorgang") = 75 'ID entspricht 'Vorgang erfasst.'
        rs("Kürzel") = fKuerzel
        rs("StdSatz") = lStdSatz
        rs("StdSatzID") = lStdSatzId
        rs("DatumStdSatz") = Date
        rs("Aufwand") = 0.25 * lStdSatz

'        Set rsKuerzel = ldb.OpenRecordset("Select Kurzzeichen from Bearbeiter where Name = '" & fKuerzel & "';")
'            If rsKuerzel.RecordCount Then
'                rs("Kürzel") = rsKuerzel("Kurzzeichen")
'                rsKuerzel.Clone
'                Set rsKuerzel = Nothing
'            End If
        rs.Update
        'Den Aufwand der Reklamationsanlage in das Feld der Summer aller Aufwände eintragen
        ldb.Execute ("Update Mastertab set [Zeitbedarf] = 0.25 where [ID] = " & fReklID & ";")
    End If
    rs.MoveLast
    If IsNull(rs("Datum")) Then
        MsgBox getTC_FCFromCode("TC0333", ldb)  'MsgBox "Kein neuer Datensatz möglich, da in der Letzten Position Datum fehlt!"   'TC0333
    Else
        sLastPos = rs("Pos")
        rs.AddNew
        rs("MID") = fReklID
        rs("Datum") = Date
        rs("Pos") = CStr(CInt(sLastPos) + 10) 'Pos um 10 erhöhen
        rs("Kürzel") = fKuerzel
        rs("Zeit") = 0.25
        rs("StdSatz") = lStdSatz
        rs("StdSatzID") = lStdSatzId
        rs("DatumStdSatz") = Date
        rs("Aufwand") = 0.25 * lStdSatz
        rs.Update
    End If
    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")
    Exit Sub
err_NewPosAufwand2:
    Resume


End Sub

Public Function getCaption() As String
    getCaption = "TestCaption"
End Function

Public Function getAnzeigefromTable(ByVal fTab As String, fControl As String, ByRef fValue As Object, fOutput As String) As String
'Funktionsbeschreibung:
'=getAnzeigefromTable("Bearbeiter";"Name";[Zust.Service-MA];"Kurzzeichen")
'
'
On Error GoTo err_getAnzeigefromTable

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

If fTab = "Baugruppen-avista_out" Then
'    fTab = fTab
End If

If fTab = "Bearbeiter" Then
'    Debug.Print fTab
End If
If IsNull(fValue) Then
    getAnzeigefromTable = ""
Else
    Set ldb = openAccDBConnection("openDB")
    Select Case VarType(fValue)
        Case 2, 3, 4, 5
err_:
            Set rs = ldb.OpenRecordset("Select [" & fOutput & "] from [" & fTab & "]  where [" & fControl & "] = " & fValue & ";")
        Case 8
            Set rs = ldb.OpenRecordset("Select [" & fOutput & "] from [" & fTab & "]  where [" & fControl & "] = '" & fValue & "';")
    End Select

    If rs.RecordCount Then
        If IsNull(rs(fOutput)) Then
            getAnzeigefromTable = ""
        Else
            getAnzeigefromTable = rs(fOutput)
        End If
    Else
        getAnzeigefromTable = ""
    End If
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

Exit Function

err_getAnzeigefromTable:
    If Not rs Is Nothing Then Set rs = Nothing
'    If Not db Is Nothing Then Set db = Nothing
'    getAnzeigefromTable = ""
    GoTo err_

End Function
Public Function getAnzeigefromTable3(ByVal fControl As String, fTab As String, fKriterium As String) As Variant
On Error GoTo err_getAnzeigefromTable3
    Dim ldb As DAO.Database
    Dim lret As Variant


    Set ldb = openAccDBConnection("openDB")

    lret = Null
    lret = Nz(DLookup(fControl, fTab, fKriterium))
    getAnzeigefromTable3 = lret

    Exit Function

err_getAnzeigefromTable3:
        If Err = 3075 Then Resume Next
End Function



Public Function getAnzeigefromTable2(ByVal fTab As String, fControl As String, ByRef fValue As String, fOutput As String, fForm As String) As String
'Funktionsbeschreibung:
'=getAnzeigefromTable2("Kategorien-Vorgangstatus";"Kategorie-Vorgangstatus";"Mastertab#Kategorie-Status";"Anzeige";"Fehlererfassung - Register")
'
'
On Error GoTo err_getAnzeigefromTable2

Dim ldb As DAO.Database
Dim lrs_FrmQuery As DAO.Recordset
Dim rs As DAO.Recordset
Dim lfrm As Form
Dim larrField() As String


Set lfrm = Forms(fForm)

Set ldb = openAccDBConnection("openDB")
Set lrs_FrmQuery = ldb.OpenRecordset(lfrm.RecordSource)
larrField = Split(fValue, "#")

If fTab = "Baugruppen-avista_out" Then
'    fTab = fTab
End If

If fTab = "Bearbeiter" Then
'    Debug.Print fTab
End If
If IsNull(fValue) Then
    getAnzeigefromTable2 = ""
Else
    Select Case VarType(fValue)
        Case 2, 3, 4, 5
            Set rs = ldb.OpenRecordset("Select [" & fOutput & "] from [" & fTab & "]  where [" & fControl & "] = " & lrs_FrmQuery(larrField(1)) & ";")
        Case 8
            Set rs = ldb.OpenRecordset("Select [" & fOutput & "] from [" & fTab & "]  where [" & fControl & "] = '" & lrs_FrmQuery(larrField(1)) & "';")
    End Select

    If rs.RecordCount Then
        If IsNull(rs(fOutput)) Then
            getAnzeigefromTable2 = ""
        Else
            getAnzeigefromTable2 = rs(fOutput)
        End If
    Else
        getAnzeigefromTable2 = ""
    End If
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

Exit Function

err_getAnzeigefromTable2:
    If Not rs Is Nothing Then Set rs = Nothing
'    If Not db Is Nothing Then Set db = Nothing
    getAnzeigefromTable2 = ""

End Function

Public Sub KeyDown(KeyCode As Integer, Shift As Integer, frm As Form, Optional fPage As Page = Nothing)
    Dim sForm As String
    Dim acType As String
    Dim sKoord As String

    'Debug.Print "KeyCode: " & KeyCode
    sForm = frm.Name
    'MsgBox "KeyDown: " & KeyCode & " - Shift: " & Shift
    Select Case KeyCode
    Case 17 'Strg
'        bKeyStrg = True
    Case 70 'F
        If Shift = 3 Then
            'MsgBox "Set Pos"
            sKoord = setFktKoordinaten(frm)
        End If
    Case 73 'I
        If Shift = 3 Then
             sKoord = initFktKoordinaten(frm)
         End If
    Case 65 'A - Öffnen Konfig Dialog
        If Shift = 3 Then
            DoCmd.OpenForm "KonfiguartionLQReklDB"
        End If
    Case 69 'E - Append Extern DB Dialog
        If Shift = 3 Then
            Append_TableDefs2
'            neueMLeiste
        End If

    Case 66 'B - Öffnen CommandButton Steuerung
        If Shift = 3 Then
            DoCmd.OpenForm "frmCommandButtonSteuerung"
            Forms("frmCommandButtonSteuerung").Init_Form_Cmb frm, frm.ActiveControl, fPage
        End If

    Case 67 'C - Öffnen Controls Steuerung
        If Shift = 3 Then
            DoCmd.OpenForm "frmControlsSteuerung"
            Forms("frmControlsSteuerung").Init_Form frm, frm.ActiveControl, fPage
        End If

    Case 80 'P - Ausgabe in debug.print aller Dictionaries
        If Shift = 3 Then
            PrintDebugAllFormFieldValues
        End If

    Case 71 'G - Control Info  20161111-ROS
        If Shift = 3 Then

            Select Case frm.ActiveControl.ControlType
                Case 108
                    acType = "BoundObjectFrame"
                Case 106
                    acType = "CheckBox"
                Case 111
                    acType = "ComboBox"
                Case 104
                    acType = "CommandButton"
                Case 119
                    acType = "CustomControl"
                Case 103
                    acType = "Image"
                Case 100
                    acType = "Label"
                Case 102
                    acType = "Line"
                Case 110
                    acType = "ListBox"
                Case 114
                    acType = "ObjectFrame"
                Case 105
                    acType = "OptionButton"
                Case 107
                    acType = "OptionGroup"
                Case 124
                    acType = "Page"
                Case 118
                    acType = "PageBreak"
                Case 101
                    acType = "Rectangle"
                Case 112
                    acType = "Subform"
                Case 123
                    acType = "TabCtl"
                Case 109
                    acType = "TextBox"
                Case 122
                    acType = "ToggleButton"
            End Select

            If frm.ActiveControl.Name <> "" Then
                If FieldExistInRecordset(frm.Recordset, frm.ActiveControl.ControlSource) <> "" Then
                    MsgBox "Formularname: " & frm.Name & vbCr & "Aktives-Textfeld-Name: " & frm.ActiveControl.Name & vbCr & "Feld-Typ: " & acType & vbCr & _
                    "Datenquelle: " & frm.Recordset.Fields(frm.ActiveControl.ControlSource).SourceTable & vbCr & "ControlSource: " & frm.Controls(frm.ActiveControl.Name).ControlSource, , "Info's zum Datenfeld"
                    KeyCode = 0
                Else
                    MsgBox "Diesem Steuerelement ist kein Feld der hinterlegten Datenstruktur zugeordnet!"
                End If
            End If
        End If

    End Select

End Sub


Public Sub KeyUp(KeyCode As Integer, Shift As Integer)  '20161111-ROS

'    If KeyCode = 17 Then
'        bKeyStrg = False
'    End If

End Sub

Public Function FieldExistInRecordset(frs As DAO.Recordset, fName As String) As String   '20161111-ROS
'Gibt bei positiver Rückmeldung den SourceTabel-Name zurück, ansonsten Leer = ""
Dim i As Integer

    FieldExistInRecordset = ""

    For i = 0 To frs.Fields.Count - 1
        If frs.Fields(i).Name = fName Then
            FieldExistInRecordset = frs.Fields(fName).SourceTable
            Exit For
        End If
        If i = frs.Fields.Count - 1 Then
'            Debug.Print i
        End If
    Next i

End Function

Function RSPDouble(s As String) As Double

    RSPDouble = CDbl(s)

End Function
Public Function fctComboCol(strForm As String, strCombo As String, intCol As Integer)
On Error GoTo err_fctComboCol

    fctComboCol = Forms(strForm).Controls(strCombo).Column(intCol)

Exit Function

err_fctComboCol:
    If Err = 2467 Then
        fprintLogRSP "AllgemeineFunktion - fctComboCol", _
                     " Form: " & strForm & " ; Combo: " & strCombo & " ; Spalte: " & intCol
        fctComboCol = Null
        Resume Next
    End If

End Function



Function getDateUSSQL(fdate As Date) As String

getDateUSSQL = "#" & format(fdate, "mm\/dd\/yyyy") & "#"

End Function

Public Sub NewReklamationE2(fReason, f_ReklNr) '20160205-ROS

Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean
Dim ReklBearbeiten As Boolean

If fReason = "Bearbeiten" Then
    If f_ReklNr = "" Then
        f_ReklNr = get_p_ReklNr()
    Else
        p_ReklNr = f_ReklNr
    End If
Else
    p_ReklNr = ""
End If

'Set db = currentdb ()
NewRekl = False
'MsgBox "Reklamation: " & ReklNr

If get_p_ReklNr() = "" Then
    'ermitteln letzte ReklNr
    p_ReklNr = getLastDSMastertab()
    If get_p_ReklNr() > 0 Then
        'ReklNr erhöhen = Neuer Vorgang
        p_ReklNr = get_p_ReklNr() + 1
        'Öffnen der Fehlererfassung mit neuer Rekl-Nr
        ' -> Notwendig damit die Reklamtionsnummer sofort generiert wird und für weiter Zwecke genutzt werden kann
        If (setNewReklDS(get_p_ReklNr(), format(Date, "dd.mm.yy"), "Bearbeitung: Service", "berechtigt", "extern") = 0) Then Exit Sub
        '20151002-RSP Filter direkt setzen
        DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit
        'Das aktuelle Datum wird als Auftragsdatum eingetragen
        'Ermittlung Zuständiger Service Mitarbeiter (Erfasser - Extern)
         Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Kurzzeichen] = '" & get_pKuerzel() & "';")
        If rs.RecordCount Then
            ldb.Execute ("Update Mastertab set [Zuständiger Service-MA] = '" & rs("Name").Value & "' where [ID] = " & get_p_ReklNr() & ";")
            NewPosAufwand get_p_ReklNr(), rs("Name").Value
        Else
            MsgBox Replace(getTC_FCFromCode("TC0434"), "<xxx>", get_pKuerzel()), vbOKOnly, getTC_FCFromCode("TC0407")
        End If
        'Template Fehlergrundbeschreibung in Bearbeiter Daten 2
        ldb.Execute ("Update Mastertab set [Wie Warum passierte der Fehler] = 'Fehlergrund:" & _
            Chr(13) & Chr(10) & _
            Chr(13) & Chr(10) & _
            "Qualifizierung: berechtigt" & _
            Chr(13) & Chr(10) & "' where [ID] = " & get_p_ReklNr() & ";")

        'Anlegen des Datensatz zum Sperren der Bearbeitung durch einen Zweiten Benutzer
        lockDataset (get_p_ReklNr())
        'ReklBearbeiten = True
        '
        'InsertNewReklMemoDS (get_get_p_ReklNr ()l ())
        'Beim Öffnen so setzen das der generierte Kommentar nicht kopiert wird. Kopie macht erst Sinn wenn Änderungen
        'am Datensatz bzw. direkt im Kommentarfeld vollzogen wurden. In dem Fall wird die Variable wie folgt geändert:
        'Bei einer Änderung im Steuerelement wird die Variable auf True gesetzt (
        p_KommentarKopieren = False

        NewRekl = True
        'Ordner für die Dokumentenablage anlegen - Ordnername = ReklNr
        With Forms("Fehlererfassung - Register")
'            If Dir("F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\" & Format(Date, "yyyy") & "\" & .ID & "\", vbDirectory) = "" Then
'                MkDir ("F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\" & Format(Date, "yyyy") & "\" & .ID & "\")
'            End If
            If Dir(getTC_FCFromCode("DP0004") & format(Date, "yyyy") & "\" & .ID & "\", vbDirectory) = "" Then
                MkDir (getTC_FCFromCode("DP0004") & format(Date, "yyyy") & "\" & .ID & "\")
            End If
        End With

        'Forms("Fehlererfassung - Register")!UF_Fehlererfassung.Requery

        'GoTo end_Befehl41_Click
    End If
Else
    'gibt es die eingegeben Reklamationsnummer?
    Dim strKrit As String
    strKrit = "ID = " & get_p_ReklNr()
    If DCount("ID", "Mastertab", strKrit) = 0 Then
        MsgBox Replace(getTC_FCFromCode("TC0440"), "<xxx>", get_p_ReklNr())    '("Reklamation mit der Nummer: " & get_get_p_ReklNr ()l () & " ist nicht existent!")
        GoTo end2_NewReklamationE2
    Else
        'Öffnen der Fehlererfassung mit vorhandener Rekl-Nr.
        '20151002-RSP Filter direkt setzen
        'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_get_p_ReklNr ()l (), acFormEdit
        If GetStateVorgang(get_p_ReklNr()) = "Abgeschlossen" Then
            DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
            'Forms("Fehlererfassung - Register").RecordsetType = 2 'Snapshot
        Else
            If proofAndLockDataset(get_p_ReklNr()) = "" Then
                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormEdit
                ReklBearbeiten = True
            Else
                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab].[ID] = " & get_p_ReklNr(), acFormReadOnly
                ReklBearbeiten = False
                'Forms("Fehlererfassung - Register").RecordsetType = 2 'Snapshot
            End If
        End If
    End If

End If
'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
If (ReklBearbeiten) Then
    InsertNewReklMemoDS (get_p_ReklNr())
    InsertNewEinkaufDS (get_p_ReklNr())
End If
'Aktiviert alle Register in der Fehlerübersicht
'set_FehlererfassungEnviroment
'###
With Forms("Fehlererfassung - Register")
    .Requery
    .Refresh
    .AllowAdditions = False
    'DoCmd.Maximize

    'Sichern des aktuellen Vorgang-Status
    If IsNull(Forms("Fehlererfassung - Register").Controls("Kategorie-Status")) Then
        p_VorgangStatus = ""
    Else
        p_VorgangStatus = Forms("Fehlererfassung - Register").Controls("Kategorie-Status")
    End If
    'Forms("Fehlererfassung - Register").Controls("Kunden-ID").SetFocus  '20121121-RSP Änderung
End With

'20160205-ROS
If IsNull(get_pKuerzel()) Then
        NewPosAufwand get_p_ReklNr(), ""
    Else
        NewPosAufwand get_p_ReklNr(), get_pKuerzel()
    End If
        With Forms("Fehlererfassung - Register")
        .Requery
        .Refresh
        End With
    DoCmd.GoToRecord , , acLast


If Not IsNull(rs) Then Set rs = Nothing
'If Not IsNull(db) Then Set db = Nothing


end_NewReklamationE2:
end2_NewReklamationE2:


End Sub

Public Sub MahnwesenKunde(ByVal fReklID As Long, fKuerzel As String) '20160302-ROS
On Error GoTo err_MahnwesenKunde

    Dim ldb As DAO.Database
    Dim rs, rsKuerzel As DAO.Recordset
    Dim sLastPos As String
    Dim DEEri As String
    Dim DEMah As String

    If (getCurrLanguage = "German") Then
        DEEri = "Bearbeitungserinnerung"
        DEMah = "Mahnung"
    Else
        DEEri = "working memory"
        DEMah = "reminder"
    End If


    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblMahnwesenKunde where MID=" & fReklID & " ORDER BY VorgangID;")
    If rs.RecordCount = 0 Then
        'Anlegen Erfassungsposition 0
        rs.AddNew
        rs("MID") = fReklID
        rs("Stufe") = "0"
        If rs("Stufe") = "0" Then
            rs("Datum") = Date
            rs("Mahnstufe") = DEEri
        End If
        rs.Update
    Else
    rs.MoveLast
        If IsNull(rs("Datum")) Then
            MsgBox getTC_FCFromCode("TC0333", ldb)  'MsgBox "Kein neuer Datensatz möglich, da in der Letzten Position Datum fehlt!"   'TC0333
        Else
            sLastPos = rs("Stufe")
            rs.AddNew
            rs("MID") = fReklID
            rs("Stufe") = CStr(CInt(sLastPos) + 1) 'Pos um 1 erhöhen
            rs("Datum") = Date
            rs("Mahnstufe") = DEMah
            rs.Update
            rs.Close
        End If
    End If
    Set rs = Nothing

    Exit Sub
err_MahnwesenKunde:
    Resume


End Sub
Public Function get_ServiceZuordnung() As Boolean  '20161129-ROS

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * FROM Bearbeiter WHERE  [Name] =  '" & fOSUserName & "' and [Aktiv] = True;", dbOpenSnapshot)
    If rs.RecordCount > 0 Then
            get_ServiceZuordnung = True
        Else
            get_ServiceZuordnung = False
    End If

    Set rs = Nothing

End Function
Public Function get_AdminZuordnung() As Boolean '20170125-ROS
On Error GoTo err_get_AdminZuordnung

Dim dbDestination As DAO.Database
Dim tdf As DAO.TableDef
Dim sDatei As String


Dim ldb As DAO.Database
Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT *FROM KonfigurationReklDB WHERE  [Konfiguration] =  '" & fOSUserName & "' and [Value] = 'Admin';")
    If rs.RecordCount > 0 Then
            get_AdminZuordnung = True
        Else
            get_AdminZuordnung = False
    End If

    Set rs = Nothing
    Exit Function

err_get_AdminZuordnung:

    If Err = 3078 Then ' Fehlerhafter Recordsetaufruf beim Ermitteln des Adminpasswort
        MsgBox "Keine Datenbank vorhanden, bitte anschließend eine Datenbank auswählen!"
        sDatei = DateiOeffnen()

        Set dbDestination = CurrentDb

        Set tdf = dbDestination.CreateTableDef("KonfigurationReklDB")
        tdf.Connect = ";DATABASE=" & sDatei
        tdf.SourceTableName = "KonfigurationReklDB"
        dbDestination.TableDefs.Append tdf
        Set tdf = Nothing
        Set tdf = dbDestination.CreateTableDef("tblMenüleiste")
        tdf.Connect = ";DATABASE=" & sDatei
        tdf.SourceTableName = "tblMenüleiste"
        dbDestination.TableDefs.Append tdf
        Set tdf = Nothing
        Set tdf = dbDestination.CreateTableDef("tblRegister")
        tdf.Connect = ";DATABASE=" & sDatei
        tdf.SourceTableName = "tblRegister"
        dbDestination.TableDefs.Append tdf
        Set tdf = Nothing

        Resume
    End If
End Function


Public Function getPWDBAdmin() As String
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'AdminPW';")
    If rs.RecordCount Then
        If IsNull(rs("ValueEncrypt")) Then
            getPWDBAdmin = ""
        Else
            getPWDBAdmin = rs("ValueEncrypt")
        End If
    End If
    Set rs = Nothing

End Function
Public Function getPWDBUser() As String      '20170125-ROS
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'UserPW';")

    If rs.RecordCount Then
        If IsNull(rs("ValueEncrypt")) Then
            getPWDBUser = ""
        Else
            getPWDBUser = rs("ValueEncrypt")
        End If
    End If
    Set rs = Nothing

End Function

Sub sendMailAufgabeAnhang(ByVal fsEmpfaenger As String, ByVal fsAbsender As String, ByVal fsBetreff As String, ByVal fsBody As String, fcolAnhang As Collection, _
                          ByVal fbMailOeffnen As Boolean, Optional ByVal fsCC As String = "", Optional ByVal fsBCC As String = "")
On Error GoTo err_sendMailLieferantAnhang  '20170328-ROS

    Dim lclsMailProv As clsMailProviderOutlook

    Dim iMsg As Object
    Dim iConf As Object
    Dim ldb As DAO.Database
    Dim Entwicklung As String
    Dim Flds
    Dim i As Integer

    Set ldb = openAccDBConnection("openDB")

    'Instanz auf einen Outlook MailProvider
    Set lclsMailProv = New clsMailProviderOutlook
    'Mailinhalte = Body füllen
    lclsMailProv.setBody = fsBody
    lclsMailProv.setBetreff = fsBetreff
    'Mailobjekt füllen, versenden und abspeichern
    lclsMailProv.setAn = fsEmpfaenger
    lclsMailProv.setVon = fsAbsender
    lclsMailProv.setCC = fsCC & "; " & fsEmpfaenger
    lclsMailProv.setBCC = fsBCC
    For i = 1 To fcolAnhang.Count
        lclsMailProv.AddItemMailItem (fcolAnhang.Item(i))
    Next i

    lclsMailProv.fillMailItem
    lclsMailProv.SaveAsMailItem get_DateiPfad & "E-Mail\", get_pKuerzel
    If DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Entwicklung'") = "Nein" Then
        DoCmd.Hourglass True
        lclsMailProv.sendMailItem
        DoCmd.Hourglass False
    End If
    Set lclsMailProv = Nothing

    Exit Sub
err_sendMailLieferantAnhang:

    MsgBox Err.Description

End Sub

Public Function HtmlTextCreate(ByVal strText As String, ByRef strTextHtml As String) '20170328-ROS

', ByRef strTextHtml As String

        strText = Replace(strText, "&", "&amp;")
        strText = Replace(strText, """", "&quot;")
'        strText = Replace(strText, "'", "''")
        strText = Replace(strText, "<", "&lt;")
        strText = Replace(strText, ">", "&gt;")
        strText = Replace(strText, vbCrLf, "<br>")
        strText = Replace(strText, "  ", " &nbsp;")
        strText = Replace(strText, Chr(13), "<br>")


        strTextHtml = "<span style=""font-size:10pt; font-family:''Arial''"">" & strText
'        HtmlTextCreate = "<span style=""font-size:10pt; font-family:''Arial''"">" & strText

End Function
Public Function ChangeSonderzeichenSQLText(ByVal strTextQ As String, ByRef strTextZ As String) '20170328-ROS

        strTextZ = Replace(strTextZ, "'", "''")

End Function
Public Function IsFormDialog(frm As Form) As Boolean
'#################################################################################################################################
'Funktion dient zur Abfrage ob ein Formular in acDialog göffnet ist
'Muss in dem abgefragten Formular hinterlegt werden:
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
'Private Const GWL_EXSTYLE As Long = -20
'Private Const WS_EX_DLGMODALFRAME As Long = &H1
'#################################################################################################################################
    Dim hwnd As LongPtr
    Dim lngStyle As Long

    hwnd = frm.hwnd
    lngStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    IsFormDialog = CBool((lngStyle And WS_EX_DLGMODALFRAME) = WS_EX_DLGMODALFRAME)

End Function


Public Function getClientShutdownOwner() As String

    If IsNull(DLookup("Value", "KonfigurationReklDB", "Konfiguration = 'ClientShutdownOwner'")) Then
        getClientShutdownOwner = ""
    Else
        getClientShutdownOwner = DLookup("Value", "KonfigurationReklDB", "Konfiguration = 'ClientShutdownOwner'")
    End If
End Function



Public Function getTicketStatusFarbe(fl_ReklID As Long, fs_Farbe As String) As String

    Dim ldb As DAO.Database
    Dim rsTicketStatus As DAO.Recordset
    Dim rs As DAO.Recordset
    Dim lsSQL1 As String
    Dim lsSQL2 As String


    lsSQL1 = "SELECT tblTicketListe.TicketID " & _
            "FROM (tblTicketListe INNER JOIN Mastertab ON tblTicketListe.ReklID = Mastertab.ID) " & _
            "INNER JOIN (Sachbearbeiter AS Sachbearbeiter_1 INNER JOIN (LQEinheit INNER JOIN Sachbearbeiter ON LQEinheit.LQEinheit = Sachbearbeiter.Einheit) " & _
            "ON Sachbearbeiter_1.Einheit = LQEinheit.LQEinheit) ON Mastertab.LQEinheitID = LQEinheit.ID " & _
            "WHERE (tblTicketListe.ReklID = " & fl_ReklID & ") And (tblTicketListe.TicketStatusID = 2) And (tblTicketListe.VerlängertBis < Date()) " & _
            "GROUP BY tblTicketListe.TicketID, Sachbearbeiter_1.Kurzzeichen " & _
            "HAVING (Sachbearbeiter_1.Kurzzeichen ='" & get_pKuerzel() & "');"

    lsSQL2 = "SELECT tblTicketListe.TicketID " & _
            "FROM (tblTicketListe INNER JOIN Mastertab ON tblTicketListe.ReklID = Mastertab.ID) " & _
            "INNER JOIN (Sachbearbeiter AS Sachbearbeiter_1 INNER JOIN (LQEinheit INNER JOIN Sachbearbeiter ON LQEinheit.LQEinheit = Sachbearbeiter.Einheit) " & _
            "ON Sachbearbeiter_1.Einheit = LQEinheit.LQEinheit) ON Mastertab.LQEinheitID = LQEinheit.ID " & _
            "WHERE (tblTicketListe.ReklID = " & fl_ReklID & ") And (tblTicketListe.TicketStatusID = 2) " & _
            "GROUP BY tblTicketListe.TicketID, Sachbearbeiter_1.Kurzzeichen " & _
            "HAVING (Sachbearbeiter_1.Kurzzeichen ='" & get_pKuerzel() & "');"


    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(lsSQL2, dbOpenSnapshot)
'    If fl_ReklID = 39857 Then
'        Debug.Print fl_ReklID
'    End If
'    Debug.Print rsTicketStatus.RecordCount
   If IsNull(rs.RecordCount) Or rs.RecordCount = 0 Then
    Exit Function
   Else
        Set rsTicketStatus = ldb.OpenRecordset(lsSQL1, dbOpenSnapshot)
        If rsTicketStatus.RecordCount Then
            getTicketStatusFarbe = 2 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Rot.jpg" 'min. ein Ticket ist rot"
        Else
            lsSQL1 = Replace(lsSQL1, "<", "=")
            Set rsTicketStatus = ldb.OpenRecordset(lsSQL1, dbOpenSnapshot)
            If rsTicketStatus.RecordCount Then
                getTicketStatusFarbe = 1 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Geld.jpg" 'min. ein Ticket ist gelb
            Else
                getTicketStatusFarbe = 0 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Grün.jpg" 'Alles Grün
            End If

        End If

    Select Case fs_Farbe

        Case "Rot"
'            If getTicketStatusFarbe = 2 Then getTicketStatusFarbe = "\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Rot.jpg" 'min. ein Ticket ist rot"
            If getTicketStatusFarbe = 2 Then getTicketStatusFarbe = getTC_FCFromCode("DP0012") 'min. ein Ticket ist rot"
        Case "Gelb"
'            If getTicketStatusFarbe = 1 Then getTicketStatusFarbe = "\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Gelb.jpg" 'min. ein Ticket ist gelb
            If getTicketStatusFarbe = 1 Then getTicketStatusFarbe = getTC_FCFromCode("DP0010") 'min. ein Ticket ist gelb
        Case "Grün"
'            If getTicketStatusFarbe = 0 Then getTicketStatusFarbe = "\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Grün.jpg" 'Alles Grün
            If getTicketStatusFarbe = 0 Then getTicketStatusFarbe = getTC_FCFromCode("DP0011") 'Alles Grün
    End Select
    End If

    Set rs = Nothing
    Set rsTicketStatus = Nothing

End Function
Public Function getTicketMahnungStatusFarbe(fl_ReklID As Long, fs_Farbe As String) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsTicketStatus As DAO.Recordset
    Dim lsSQL0 As String
    Dim lsSQL1 As String
    Dim lsSQL2 As String
    Dim lMaßnahmeErledigt As String

    Set ldb = openAccDBConnection("openDB")

    lMaßnahmeErledigt = Nz(DLookup("[UmgesetztAm]", "Maßnahmen", "[MaßnahmeID] = " & fl_ReklID & ""))

    Set rs = ldb.OpenRecordset("SELECT tblTicketListe.ReklID FROM tblTicketListe WHERE (tblTicketListe.ReklID= " & fl_ReklID & ") AND ((tblTicketListe.TicketStatusID = 2) or (tblTicketListe.TicketStatusID = 3));")
            If rs.RecordCount Then

                lsSQL2 = "SELECT Maßnahmen.MaßnahmeID " & _
                        "From Maßnahmen " & _
                        "WHERE (Maßnahmen.UmgesetztAm Is Null) AND (Maßnahmen.EndeUmsetzung < Date()) AND (Maßnahmen.MaßnahmeID= " & fl_ReklID & ");"

                lsSQL0 = "SELECT Maßnahmen.MaßnahmeID " & _
                        "From Maßnahmen " & _
                        "WHERE (Maßnahmen.UmgesetztAm Is not Null) AND (Maßnahmen.MaßnahmeID= " & fl_ReklID & ");"

                lsSQL1 = "SELECT Maßnahmen.MaßnahmeID " & _
                        "From Maßnahmen " & _
                        "WHERE (Maßnahmen.UmgesetztAm Is Null) AND (Maßnahmen.MaßnahmeID= " & fl_ReklID & ")and ((Maßnahmen.EndeUmsetzung Is Null) or (Maßnahmen.EndeUmsetzung >= Date()));"

                Set rsTicketStatus = ldb.OpenRecordset(lsSQL2, dbOpenSnapshot)
                If rsTicketStatus.RecordCount Then
                    getTicketMahnungStatusFarbe = 2 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Rot.jpg" 'min. ein Ticket ist rot"
                Else
                    Set rsTicketStatus = ldb.OpenRecordset(lsSQL1, dbOpenSnapshot)
                    If rsTicketStatus.RecordCount Then
                        getTicketMahnungStatusFarbe = 1 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Geld.jpg" 'min. ein Ticket ist gelb
                    Else
                        Set rsTicketStatus = ldb.OpenRecordset(lsSQL0, dbOpenSnapshot)
                        If rsTicketStatus.RecordCount Then
                            getTicketMahnungStatusFarbe = 0 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Grün.jpg" 'Alles Grün
                        End If
                    End If

                End If

                Select Case fs_Farbe

                    Case "Rot"
                        If getTicketMahnungStatusFarbe = 2 Then getTicketMahnungStatusFarbe = getTC_FCFromCode("DP0012") 'min. ein Ticket ist rot"
                    Case "Gelb"
                        If getTicketMahnungStatusFarbe = 1 Then getTicketMahnungStatusFarbe = getTC_FCFromCode("DP0010") 'min. ein Ticket ist gelb
                    Case "Grün"
                        If getTicketMahnungStatusFarbe = 0 Then getTicketMahnungStatusFarbe = getTC_FCFromCode("DP0011") 'Alles Grün
                End Select

                Set rsTicketStatus = Nothing
                Set rs = Nothing
            Else
                '20260305-ROS für Maßnahmen die kein Ticket haben und abgeschlossen sind wir der Status Grün gesetzt
                If lMaßnahmeErledigt <> "" And fs_Farbe = "Grün" Then
                    getTicketMahnungStatusFarbe = getTC_FCFromCode("DP0011")
                    Set rs = Nothing
                Else
                    Set rs = Nothing
                    Exit Function
                End If
            End If

End Function

Public Function TicketMWAnlegen(frReklID As Long, frTicketTypID As Long, frMailDatum As Date, frBearbeitungBis As Date, frVerlängertBis As Date, _
                                frTicketStatusID As Long, frMailErstellt As Date, frMailVon As String, frMailTo As String, frMailCC As String, _
                                frMailBCC As String, Optional frTelefon As String) As Long

        Dim ldb As DAO.Database
        Dim rs As Recordset
        Dim rs2 As Recordset
        Dim rs3 As Recordset
        Dim strTextHtml As String

        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from tblTicketListe where [ReklID] = " & frReklID & " and [TicketTypID] = " & frTicketTypID & " ;")
'        Set rs2 = ldb.OpenRecordset("Select *from qryTicketMaßnahmenPE where [SB-ID] =  " & frKuerzel & ";")
        Set rs3 = ldb.OpenRecordset("Select *from tblTicketTyp where [TicketTypID] =  " & frTicketTypID & ";")

        If rs.RecordCount Then
'            MsgBox ("Mahnwesen wurde schon aktviert")
            MsgBox getTC_FCFromCode("TC3191"), vbInformation, getTC_FCFromCode("TC3192")
            Exit Function
        Else
            rs.AddNew
            rs("ReklID") = frReklID
            rs("TicketTypID") = frTicketTypID
            rs("MailVon") = frMailVon
            rs("MailAN") = frMailTo
            rs("MailCC") = frMailCC
            rs("MailBCC") = frMailBCC
            If rs3.RecordCount Then
                rs("MailBetreff") = rs3("MailBetreff")
                rs("MailBetreff") = Replace(rs3("MailBetreff"), "<ReklNr>", frReklID)
                rs("MailText") = rs3("MailText")
            Else
                rs("MailText") = ""
                rs("MailBetreff") = ""
            End If
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", frReklID)
            rs("MailText") = Replace(rs("MailText"), "<Bis>", frBearbeitungBis)
            rs("MailText") = Replace(rs("MailText"), "<E-Mail_Absender>", frMailVon)
            rs("MailText") = Replace(rs("MailText"), "<Telefon>", frTelefon)
            HtmlTextCreate rs("MailText"), strTextHtml
            rs("MailTextHTML") = strTextHtml
            rs("MailDatum") = frMailDatum
            rs("BearbeitungBis") = frBearbeitungBis
            rs("VerlängertBis") = frVerlängertBis
            rs("TicketStatusID") = frTicketStatusID
            rs("MailErstellt") = frMailErstellt
            TicketMWAnlegen = rs("TicketID")
            rs.Update
        End If

        Set rs = Nothing
'        Set rs2 = Nothing
        Set rs3 = Nothing

End Function

Public Function DatumAenderungMW(frReklID As Long, frTicketTypID As Long, frMailDatum As Date, frVerlängertBis As Date, Optional frErledigtAm As String = "31.12.9999", Optional frErledigtVom As String) As String

        Dim ldb As DAO.Database
        Dim rs As Recordset



        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from tblTicketListe where [ReklID] = " & frReklID & " and [TicketTypID] = " & frTicketTypID & ";")

        If rs.RecordCount Then
            If Not IsNull(rs("ErledigtAm")) Then
                If frTicketTypID = 11 Then
'                    MsgBox ("Maßnahme ist schon umgesetzt! Keine Datenänderung mehr möglich!")
                    MsgBox getTC_FCFromCode("TC3193"), vbInformation, getTC_FCFromCode("TC3194")
                Else
'                    MsgBox ("Wirksamkeitsprüfung ist schon abgeschlossen! Keine Datenänderung mehr möglich!")
                    MsgBox getTC_FCFromCode("TC3195"), vbInformation, getTC_FCFromCode("TC3196")
                End If
                If rs("ErledigtAm") <> frErledigtAm Then DatumAenderungMW = rs("ErledigtAm")
                If rs("MailDatum") <> frMailDatum Then DatumAenderungMW = rs("MailDatum")
                If rs("VerlängertBis") <> frVerlängertBis Then DatumAenderungMW = rs("VerlängertBis")
            Else
                rs.Edit
                If frErledigtAm < "31.12.9999" Then
                    rs("ErledigtAm") = frErledigtAm
                    rs("TicketStatusID") = 3
                    rs("ErledigtVon") = frErledigtVom
                Else
                    rs("MailDatum") = frMailDatum
                    rs("VerlängertBis") = frVerlängertBis
                End If
                rs.Update
                DatumAenderungMW = "31.12.9999"
            End If
        Else
            DatumAenderungMW = "31.12.9999"
        End If

        Set rs = Nothing

End Function
Public Function getTicketlisteStatusFarbe(fl_TicketID As Long, fs_Farbe As String) As String
On Error GoTo Fehler

    Dim ldb As DAO.Database
    Dim rsTicketStatus As DAO.Recordset
    Dim lsSQL0 As String
    Dim lsSQL1 As String
    Dim lsSQL2 As String


    Set ldb = openAccDBConnection("openDB")

                lsSQL0 = "SELECT tblTicketListe.TicketID " & _
                        "From tblTicketListe " & _
                        "WHERE (((tblTicketListe.TicketID)=" & fl_TicketID & ") AND ((tblTicketListe.TicketStatusID)=1 Or (tblTicketListe.TicketStatusID)=2 Or (tblTicketListe.TicketStatusID)=3));"

                lsSQL1 = "SELECT tblTicketListe.TicketID " & _
                        "From tblTicketListe " & _
                        "WHERE (((tblTicketListe.TicketID)=" & fl_TicketID & ") AND ((tblTicketListe.TicketStatusID)=4));"

                lsSQL2 = "SELECT tblTicketListe.TicketID " & _
                        "From tblTicketListe " & _
                        "WHERE (((tblTicketListe.TicketID)=" & fl_TicketID & ") AND ((tblTicketListe.TicketStatusID)=5 Or (tblTicketListe.TicketStatusID)=6));"


                Set rsTicketStatus = ldb.OpenRecordset(lsSQL0, dbOpenSnapshot)
                If rsTicketStatus.RecordCount Then
                    getTicketlisteStatusFarbe = 0
                Else
                    Set rsTicketStatus = ldb.OpenRecordset(lsSQL1, dbOpenSnapshot)
                    If rsTicketStatus.RecordCount Then
                        getTicketlisteStatusFarbe = 1
                    Else
                        Set rsTicketStatus = ldb.OpenRecordset(lsSQL2, dbOpenSnapshot)
                        If rsTicketStatus.RecordCount Then
                            getTicketlisteStatusFarbe = 2
                        End If
                    End If
                End If


                Select Case fs_Farbe

                    Case "Rot"
                        If getTicketlisteStatusFarbe = 2 Then getTicketlisteStatusFarbe = getTC_FCFromCode("DP0012") 'min. ein Ticket ist rot"
                    Case "Gelb"
                        If getTicketlisteStatusFarbe = 1 Then getTicketlisteStatusFarbe = getTC_FCFromCode("DP0010") 'min. ein Ticket ist gelb
                    Case "Grün"
                        If getTicketlisteStatusFarbe = 0 Then getTicketlisteStatusFarbe = getTC_FCFromCode("DP0011") 'Alles Grün

                End Select
                Set rsTicketStatus = Nothing

Fehler:
                Set rsTicketStatus = Nothing

                Exit Function

End Function

Public Sub DialCOM(strNumber As String)
    ' MsgBox "DialCOM() " & strNumber
    On Error GoTo ErrorHandler

    Dim oCtiServices As Object
    Set oCtiServices = GetObject("", "ECtiClient.CtiServices")
    oCtiServices.CtiMakeCall strNumber
    ' oCtiServices.CtiMakeCall "138"
    Exit Sub

'    Dial (strNumber)


ErrorHandler:
    MsgBox "DialCOM(" & strNumber & ") failed"

End Sub
Public Sub AnsichtFehlerpositionen()

    If IsNull(Forms("Fehlererfassung - Register").Controls("Mastertab#Referenznummer")) Then
        Forms("Fehlererfassung - Register").Controls("Bezeichnungsfeld6119").Visible = False
        Forms("Fehlererfassung - Register").Controls("bez_AnzahlDerFehlerpositionen").Visible = False
        Forms("Fehlererfassung - Register").Controls("BezAnzahlFehlerpositionen").Visible = True
        Forms("Fehlererfassung - Register").Controls("BezAnzahlFehlerpositionen2").Visible = True
    Else
        Forms("Fehlererfassung - Register").Controls("Bezeichnungsfeld6119").Visible = True
        Forms("Fehlererfassung - Register").Controls("bez_AnzahlDerFehlerpositionen").Visible = True
        Forms("Fehlererfassung - Register").Controls("BezAnzahlFehlerpositionen").Visible = False
        Forms("Fehlererfassung - Register").Controls("BezAnzahlFehlerpositionen2").Visible = False
    End If

End Sub

Public Sub AendernRef()

    Dim ldb As DAO.Database
    Dim rsRef As DAO.Recordset
    Dim rsAllRef As DAO.Recordset
    Dim rsRefBasis As DAO.Recordset
    Dim la_MldgKnd() As String
    Dim lb_BasisAnzRefSetzen As Boolean



    Set ldb = openAccDBConnection("openDB")
    'Alle Referenznummern in Mastertab ermitteln
    Set rsRef = ldb.OpenRecordset("Select Referenznummer from Mastertab where not Referenznummer is null group by Referenznummer;")
    'Über alle Referenznummern
    rsRef.MoveFirst
    While Not rsRef.EOF
'        Debug.Print rsRef("Referenznummer")
        If rsRef("Referenznummer") = "FN118048_V3601511" Then
'            Debug.Print rsRef("Referenznummer")
        End If
        'Ermiiteln aller Vorgänge zu einer Referenznummer
        Set rsAllRef = ldb.OpenRecordset("Select * from Mastertab where Referenznummer = '" & rsRef("Referenznummer") & "' order by [Meldungsnummer Kunde];")
        'Ermitteln BasisVorgang zu der Referenz
        Set rsRefBasis = ldb.OpenRecordset("Select * from Mastertab where [Meldungsnummer Kunde] = '" & rsRef("Referenznummer") & "';")
        'Nur wenn Basis vorhanden können die zugehörigen Refrenzen abgeglichen werden
        If (rsRefBasis.RecordCount > 0) And (rsAllRef.RecordCount > 0) Then
            'Basis Fehlerpos Anzhal neu berechnen und setzen erst einmal nein
            lb_BasisAnzRefSetzen = False
            'Über alle ermittelten Vorgänge des der aktuellen Referenz
            'Überprüfen ob Vorgang schon korrekt umgesetzt ist
            If rsAllRef("Meldungsnummer Kunde") <> rsRefBasis("Meldungsnummer Kunde") Then
                While Not rsAllRef.EOF
                    'Ermitteln Fehlerposition des abhängigen Vorgang
                    la_MldgKnd = Split(rsAllRef("Meldungsnummer Kunde"), "_")
                    If UBound(la_MldgKnd) > 0 Then
                        lb_BasisAnzRefSetzen = True
                        'Eintrag aktuelle Position des Vorgang
                        rsAllRef.Edit
                        'In Positionsvorgang Refrenz (ReklNr Basis) eintragen
                        rsAllRef("Referenznummer") = rsRefBasis("ID")
                        rsAllRef("Fehlerpositionen") = la_MldgKnd(1)
                        'Medungsnummer Kunde aus Basis übernehmen
                        rsAllRef("Meldungsnummer Kunde") = rsRefBasis("Meldungsnummer Kunde")
                        rsAllRef.Update
                    End If
                    rsAllRef.MoveNext
                Wend
                If lb_BasisAnzRefSetzen Then
                    rsRefBasis.Edit
                    'In Basis Vorgang Anzahl Vorgänge(Psoitionen eintragen
                    rsRefBasis("Fehlerpositionen") = rsRefBasis.RecordCount - 1
                    rsRefBasis.Update
                End If
            End If
        End If
        rsRef.MoveNext
    Wend
    Set rsRefBasis = Nothing
    Set rsRef = Nothing

End Sub
Public Sub KundnRef()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Kundenref As String
    Dim LenKundenref As Integer

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Abfrage1  where not Referenznummer is null;")
        rs.MoveFirst
        While Not rs.EOF
        Kundenref = rs("Referenznummer")
        LenKundenref = Len(Kundenref)
            If LenKundenref > 5 Then
                rs.Edit
                rs("KundenRefNummer") = rs("Referenznummer")
                rs("Referenznummer") = Null
                rs.Update
            End If
            rs.MoveNext
        Wend
    Set rs = Nothing
End Sub
Public Function getMahnstufeTextTicket(fText As Long) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs1 As DAO.Recordset
    Dim rs2 As DAO.Recordset
    Dim rs3 As DAO.Recordset
    Dim rs4 As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")

    Set rs = ldb.OpenRecordset("Select * from qryOffeneAbgelaufeneTicket where TicketID = " & fText & " ;", dbOpenSnapshot)
    Set rs1 = ldb.OpenRecordset("Select * from tblTicketMahnungStatus where TicketMahnungStatusID = 1 ;", dbOpenSnapshot)    'Erinnerung
    Set rs2 = ldb.OpenRecordset("Select * from tblTicketMahnungStatus where TicketMahnungStatusID = 2 ;", dbOpenSnapshot)    '1Mahnung
    Set rs3 = ldb.OpenRecordset("Select * from tblTicketMahnungStatus where TicketMahnungStatusID = 3 ;", dbOpenSnapshot)    '2Mahnung
    Set rs4 = ldb.OpenRecordset("Select * from tblTicketMahnungStatus where TicketMahnungStatusID = 4 ;", dbOpenSnapshot)    'in Bearbeitung

        If rs.RecordCount Then
            rs.MoveLast
            If rs("Tage") = 0 Then getMahnstufeTextTicket = rs1("Anzeige")
            If rs("Tage") = 1 Then getMahnstufeTextTicket = rs1("Anzeige")
            If rs("Tage") = 2 Then getMahnstufeTextTicket = rs2("Anzeige")
            If rs("Tage") = 3 Then getMahnstufeTextTicket = rs3("Anzeige")
            If rs("Tage") > 3 Then getMahnstufeTextTicket = rs3("Anzeige")
        Else
            getMahnstufeTextTicket = rs4("Anzeige")
        End If

    Set rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
    Set rs4 = Nothing

End Function

Function fct_changedateinTicketlist(Optional fTicketID As Long = 0)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim iPos As Integer
    Dim sLeftStr As String
    Dim sRightStr As String

    Set ldb = openAccDBConnection("openDB")

    If fTicketID = 0 Then
        Set rs = ldb.OpenRecordset("Select * from tblTicketListe;")
    Else
        Set rs = ldb.OpenRecordset("Select * from tblTicketListe where TicketID = " & fTicketID & ";")
    End If

    rs.MoveFirst
    While Not rs.EOF
        If Not IsNull(rs("MailTextHTML")) And Not IsNull(rs("MailText")) And Not IsNull(rs("BearbeitungBis")) And Not IsNull(rs("VerlängertBis")) Then
            rs.Edit
            rs("MailTextHTML") = Replace(rs("MailTextHTML"), gVerlängertBis, rs("VerlängertBis"))
            rs("MailText") = Replace(rs("MailText"), gVerlängertBis, rs("VerlängertBis"))
            rs("TicketStatusID") = 2
            rs.Update
        End If
        rs.MoveNext
    Wend
    Set rs = Nothing

End Function
Public Function getNotizStatusFarbe(fl_NotizID As Long, fs_Farbe As String) As String
On Error GoTo Fehler

    Dim ldb As DAO.Database
    Dim rsNotizStatus As DAO.Recordset
    Dim lsSQL0 As String
    Dim lsSQL1 As String
    Dim lsSQL2 As String


    Set ldb = openAccDBConnection("openDB")

                lsSQL0 = "SELECT tblNotizen.NotizID " & _
                        "From tblNotizen " & _
                        "WHERE ((tblNotizen.ErledigenBis)<Date())AND tblNotizen.NotizID= " & fl_NotizID & ";"

                lsSQL2 = "SELECT tblNotizen.NotizID " & _
                        "From tblNotizen " & _
                        "WHERE ((tblNotizen.ErledigenBis)>Date())AND tblNotizen.NotizID= " & fl_NotizID & ";"

                lsSQL1 = "SELECT tblNotizen.NotizID,tblNotizen.Erledigt " & _
                        "From tblNotizen " & _
                        "WHERE ((tblNotizen.Erledigt)= True) AND tblNotizen.NotizID= " & fl_NotizID & ";"

                Set rsNotizStatus = ldb.OpenRecordset(lsSQL1, dbOpenSnapshot)
                If rsNotizStatus.RecordCount Then
                    getNotizStatusFarbe = 0 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Rot.jpg" 'min. ein Ticket ist grün"
                Else
                    Set rsNotizStatus = ldb.OpenRecordset(lsSQL0, dbOpenSnapshot)
                    If rsNotizStatus.RecordCount Then
                        getNotizStatusFarbe = 2 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Geld.jpg" 'min. ein Ticket ist rot
                    Else
                        Set rsNotizStatus = ldb.OpenRecordset(lsSQL2, dbOpenSnapshot)
                        If rsNotizStatus.RecordCount Then
                            getNotizStatusFarbe = 1 '"\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Bilder_Ampel\Ampel_Grün.jpg" 'Alles gelb
                        End If
                    End If
                End If

                Select Case fs_Farbe

                    Case "Rot"
                        If getNotizStatusFarbe = 2 Then getNotizStatusFarbe = getTC_FCFromCode("DP0012") 'min. ein Ticket ist rot"
                    Case "Gelb"
                        If getNotizStatusFarbe = 1 Then getNotizStatusFarbe = getTC_FCFromCode("DP0010") 'min. ein Ticket ist gelb
                    Case "Grün"
                        If getNotizStatusFarbe = 0 Then getNotizStatusFarbe = getTC_FCFromCode("DP0011") 'Alles Grün

                End Select
                Set rsNotizStatus = Nothing

Fehler:
                Set rsNotizStatus = Nothing

                Exit Function

End Function

Public Function GetAccountByEmail(ByRef outlApp As Outlook.Application, ByVal strSenderEmail As String) As Outlook.Account

    Dim acc     As Outlook.Account
    Dim Retval  As Outlook.Account


    For Each acc In outlApp.Session.Accounts
        If acc.SmtpAddress = strSenderEmail Then
            Set Retval = acc
            Exit For
        End If
    Next acc

    Set GetAccountByEmail = Retval

End Function


Public Sub get_KuerzelUeberWinAnmeldung()  '20180119-ROS

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim WinAnmeldung As String

    WinAnmeldung = fOSUserName
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Name] = '" & WinAnmeldung & "';", dbOpenSnapshot)
        If rs.RecordCount = 0 Or IsNull(rs("Kurzzeichen")) Then
            get_pKuerzel
        End If
    If IsNull(TempVars(cVAR_Kuerzel)) Then
        While IsNull(TempVars(cVAR_Kuerzel))
            TempVars.Add cVAR_Kuerzel, (rs("Kurzzeichen"))
        Wend
    End If
    Set rs = Nothing
End Sub
Public Function get_LoggingClient() As String  '20180124-ROS

Dim ldb As DAO.Database
Dim rs As DAO.Recordset


    If IsNull(TempVars(cVAR_LoggingClient)) Then
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDBLocal where [Konfiguration] = 'LoggingClient';", dbOpenSnapshot)
        gLoggingClient = rs("Value")
        TempVars.Add cVAR_LoggingClient, gLoggingClient
    End If

    get_LoggingClient = TempVars(cVAR_LoggingClient)

    Set rs = Nothing


End Function
Public Function get_LoggingClientAll() As String  '20180124-ROS

Dim ldb As DAO.Database
Dim rs As DAO.Recordset


    If IsNull(TempVars(cVAR_LoggingClientAll)) Then
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where [Konfiguration] = 'LoggingClientAll';", dbOpenSnapshot)
        gLoggingClientsAll = rs("Value")
        TempVars.Add cVAR_LoggingClientAll, gLoggingClientsAll
    End If

    get_LoggingClientAll = TempVars(cVAR_LoggingClientAll)

    Set rs = Nothing

End Function

Public Function get_SachbearbeiterID() As String   '20180905-ROS

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Kurzzeichen] = '" & get_pKuerzel() & "' And [Aktiv] = True;")

    get_SachbearbeiterID = CStr(rs("SB-ID"))

    'Set rs = Nothing

End Function
Public Function Aendern_pKuerzel() As String  '20180905-ROS


            TempVars.Add cVAR_Kuerzel, getValueFromInputBox(getTC_FCFromCode("TC0211"), getTC_FCFromCode("IC0219"))

                If IsNumeric(TempVars(cVAR_Kuerzel)) Or (TempVars(cVAR_Kuerzel)) = "" Then
'                        MsgBox ("Nur Buchstabeneingabe möglich")
                        MsgBox getTC_FCFromCode("TC3187"), vbInformation, getTC_FCFromCode("TC3188")
                        TempVars.Remove (cVAR_Kuerzel)
                        get_pKuerzel
                End If

    Aendern_pKuerzel = StrConv(TempVars(cVAR_Kuerzel), 1)


End Function

Public Function get_DateiPfad(Optional fReklID As Long, Optional fMelungNrKD As String, Optional fAuftragsDatum As Date) As String

    Dim sPfadDatei As String

    If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded = False Then
        If fMelungNrKD = "" Then
            sPfadDatei = getReklVerzeichnis(fReklID, format(fAuftragsDatum, "yyyy"), "")
        Else
            sPfadDatei = getReklVerzeichnis(fReklID, format(fAuftragsDatum, "yyyy"), fMelungNrKD)
        End If
    Else
        With Forms("Fehlererfassung - Register")
            If IsNull(.Controls("Mastertab#Meldungsnummer Kunde#2")) Or .Controls("Mastertab#Meldungsnummer Kunde#2") = "" Then
                sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")
            Else
                sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), .Controls("Mastertab#Meldungsnummer Kunde#2"))
            End If
        End With
    End If

    get_DateiPfad = sPfadDatei

End Function

Public Sub EMailSendenUndSpeichern(strVon As String, strKontakt As String, strBetreff As String, strBody As String, Optional sPfadDatei As String, Optional strCCKontakt As String, Optional Anhang As String)

'    Dim myOutlApp As Object
'    Dim myMail As Object
    Dim Entwicklung As String
    Dim lclsMailProv As clsMailProviderOutlook

'    Set myOutlApp = CreateObject("Outlook.Application")
'    Set myMail = myOutlApp.CreateItem(olMailItem)
    Set lclsMailProv = New clsMailProviderOutlook

'    With myMail
'        .To = strKontakt
        lclsMailProv.setAn = strKontakt
'        .CC = strCCKontakt
        lclsMailProv.setCC = strCCKontakt
'        .SentOnBehalfOfName = strVon
        lclsMailProv.setVon = strVon
'        .subject = strBetreff
        lclsMailProv.setBetreff = strBetreff
'        .HTMLBody = strBody
        lclsMailProv.setBody = strBody
'        If Anhang <> "" Then .Attachments.Add Anhang
        If Anhang <> "" Then lclsMailProv.AddItemMailItem Anhang
        lclsMailProv.fillMailItem
'        If sPfadDatei <> "" Then .SaveAs sPfadDatei & Format(Date, "yyyymmdd") & "-" & get_pKuerzel & "_" & strBetreff & ".msg"
        If sPfadDatei <> "" Then lclsMailProv.SaveAsMailItem sPfadDatei, get_pKuerzel
        Entwicklung = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Entwicklung'")
        If Entwicklung = "Ja" Then
            lclsMailProv.displayMailItem
        Else
            lclsMailProv.sendMailItem
        End If
'    End With

'    MsgBox getTC_FCFromCode("TC0428")

    'Instanz wieder freigeben
    Set lclsMailProv = Nothing

'    Set myMail = Nothing
'    Set myOutlApp = Nothing


End Sub
Public Function set_SplittenAktuellerEintrag(strTest As String)
On Error Resume Next
'    Dim strTest As String
    Dim Arr(2) As Variant

'    strTest = "21.09.2018 12:00:28 - AHÜ:"

    Arr(0) = Left(strTest, InStr(strTest, " ") - 1)
    strTest = Right(strTest, Len(strTest) - InStr(strTest, " "))
    Arr(1) = Left(strTest, InStr(strTest, " ") - 1)
    strTest = Right(strTest, Len(strTest) - InStr(strTest, " "))
    Arr(2) = strTest

    set_SplittenAktuellerEintrag = Arr(1)

End Function
Public Function get_ClientVersion() As String       '20181203-ROS

    Dim lrs As DAO.Recordset
    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    Set lrs = ldb.OpenRecordset("select * from qryVersionierungClientSortiert;")
    lrs.MoveLast
    get_ClientVersion = lrs("Version")


    'get_ClientVersion = DLast("[Datum]", "qryVersionierungClientSortiert")

    Set lrs = Nothing

End Function
Public Function get_AkuelleClientVersion() As String        '20181203-ROS


    get_AkuelleClientVersion = DLookup("Value", "KonfigurationReklDB", "Konfiguration = 'ClientVersion'")

End Function

Public Sub ClientParameterAbfragen()        '20181203-ROS

Dim Verzeichnis As String
Dim Entwicklung As String
Dim Location As String
Dim Batch_Pfad As String

    Batch_Pfad = "\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Client_Copy_1.bat"

    'Abfrage ob auf der Entwicklungsumgebung gearbeitet wird
    Entwicklung = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Entwicklung'")
    If Entwicklung = "Nein" Then
        Verzeichnis = Left(CurrentDb.Name, 1)
        'Wenn sich der Client nicht auf dem Lokalen-Laufwerk befindet, wir die Anwendung beendet
        If Verzeichnis <> "C" Then
            MsgBox (getTC_FCFromCode("TC0488")), vbOKOnly, getTC_FCFromCode("TC0486")     'Warnung!
            DoCmd.Quit
        End If
        'Abfrage der Lokalisierung
        Location = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Location'")
        If Location = "German" Then
            'Abfrage aktueller Client ->wenn DE wir der Client beendet neu kopiert und wieder gestartet         20190514-ROS
            'wenn nicht DE erfolt nur ein Warnhinweis
            If get_AkuelleClientVersion() > get_ClientVersion() Then
                MsgBox "Sie arbeiten mit einer alten Client-Version. Applikation wird beendet, aktueller Client wird kopiert und neu gestartet!"
                shell Batch_Pfad, 1
                DoCmd.Quit
            End If
        Else
            If get_AkuelleClientVersion() > get_ClientVersion() Then MsgBox (getTC_FCFromCode("TC0487") & " " & get_AkuelleClientVersion()), vbOKOnly, getTC_FCFromCode("TC0486")   'Warnung!
        End If
    End If
End Sub
Public Sub ClientAbmelden()     '20181203-ROS

Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Update tblHost set [Angemeldet] = False where HostName = '" & Environ("COMPUTERNAME") & "' and ([KZ] = get_pKuerzel()) and ([Angemeldet] = True);")

End Sub
Public Sub ClientAnmelden()     '20181203-ROS

Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Insert into tblHost ([HostName],[KZ],[Datum],[Angemeldet],[ClientVersion])values ('" & Environ("COMPUTERNAME") & "', get_pKuerzel(), Date(), True, get_ClientVersion());")

    If ldb.Name = "C:\Users\Roberto Santarelli\Documents\QS-Datenbank\Datenbank\QS-Fehlererfassung_Access_2019_64Bit_Lokal.mdb" Then Exit Sub

    If get_pKuerzel = "ROS" Or get_pKuerzel = "RSP" Then
        DoCmd.OpenForm ("frmFormularROS")
        Forms("TB:autostart").SetFocus
    End If

End Sub

Public Function get_EMailAdressen(ByRef MailAn, Optional ByRef MailCC, Optional ByRef MailBCC) As String        '20181213-ROS

    Dim ldb As DAO.Database
    Dim rsMail As DAO.Recordset


    Set ldb = openAccDBConnection("openDB")
'    Set rsMail = ldb.OpenRecordset("Select * from Mastertab where [ID] = " & Forms("Fehlererfassung - Register")![ID] & ";", dbOpenSnapshot)
    Set rsMail = ldb.OpenRecordset("Select * from Mastertab where [ID] = " & get_p_ReklNr() & ";", dbOpenSnapshot)

        If IsNull(rsMail("LQWorkstationID")) Or rsMail("LQWorkstationID") = (-1) Then
                If IsNull(rsMail("LQBereichID")) Or rsMail("LQBereichID") = (-1) Then
'                    MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & Forms("Fehlererfassung - Register")!LQEinheitID & "")
'                    MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & Forms("Fehlererfassung - Register")!LQEinheitID & "")
'                    MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & Forms("Fehlererfassung - Register")!LQEinheitID & "")
                    MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & rsMail("LQEinheitID") & "")
                    MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & rsMail("LQEinheitID") & "")
                    MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & rsMail("LQEinheitID") & "")
                Else
                    MailAn = DLookup("MailTOLQBereich", "LQBereich", "[ID] = " & rsMail("LQBereichID") & "")
                    MailCC = DLookup("MailCCLQBereich", "LQBereich", "[ID] = " & rsMail("LQBereichID") & "")
                    MailBCC = DLookup("MailEskalationLQBereich", "LQBereich", "[ID] = " & rsMail("LQBereichID") & "")
                End If
        Else
            MailAn = Nz(DLookup("MailTOLQWS", "LQWorkstation", "[ID] = " & rsMail("LQWorkstationID") & ""))
            MailCC = Nz(DLookup("MailCCLQWS", "LQWorkstation", "[ID] = " & rsMail("LQWorkstationID") & ""))
            MailBCC = Nz(DLookup("MailEskalationLQWS", "LQWorkstation", "[ID] = " & rsMail("LQWorkstationID") & ""))
        End If

        If IsNull(MailAn) Or ((MailAn) = "") Then
            MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & rsMail("LQEinheitID") & "")
            MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & rsMail("LQEinheitID") & "")
            MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & rsMail("LQEinheitID") & "")
        End If


    Set rsMail = Nothing

    get_EMailAdressen = ""

End Function

Public Function get_EMailAdressen3(ByRef MailAn, Optional ByRef MailCC, Optional ByRef MailBCC) As String       '20181217-ROS

        If IsNull(Forms("frmLagerPrüfung")![LQWorkstationID]) Or (Forms("frmLagerPrüfung")![LQWorkstationID]) = (-1) Then
                If IsNull(Forms("frmLagerPrüfung")![LQBereichID]) Or (Forms("frmLagerPrüfung")![LQBereichID]) = (-1) Then
                    MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & Forms("frmLagerPrüfung")!LQEinheitID & "")
                    MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & Forms("frmLagerPrüfung")!LQEinheitID & "")
                    MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & Forms("frmLagerPrüfung")!LQEinheitID & "")
                Else
                    MailAn = DLookup("MailTOLQBereich", "LQBereich", "[ID] = " & Forms("frmLagerPrüfung")!LQBereichID & "")
                    MailCC = DLookup("MailCCLQBereich", "LQBereich", "[ID] = " & Forms("frmLagerPrüfung")!LQBereichID & "")
                    MailBCC = DLookup("MailEskalationLQBereich", "LQBereich", "[ID] = " & Forms("frmLagerPrüfung")!LQBereichID & "")
                End If
        Else
            MailAn = Nz(DLookup("MailTOLQWS", "LQWorkstation", "[ID] = " & Forms("frmLagerPrüfung")!LQWorkstationID & ""))
            MailCC = Nz(DLookup("MailCCLQWS", "LQWorkstation", "[ID] = " & Forms("frmLagerPrüfung")!LQWorkstationID & ""))
            MailBCC = Nz(DLookup("MailEskalationLQWS", "LQWorkstation", "[ID] = " & Forms("frmLagerPrüfung")!LQWorkstationID & ""))
        End If

        If IsNull(MailAn) Or ((MailAn) = "") Then
            MailAn = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & Forms("frmLagerPrüfung")!LQEinheitID & "")
            MailCC = DLookup("MailCCLQEinheit", "LQEinheit", "[ID] = " & Forms("frmLagerPrüfung")!LQEinheitID & "")
            MailBCC = DLookup("MailEskalationLQEinheit", "LQEinheit", "[ID] = " & Forms("frmLagerPrüfung")!LQEinheitID & "")
        End If

    get_EMailAdressen3 = ""

End Function
Public Function get_IntegerfromAccessTrueFalse(Feldname As Boolean) As Integer

    get_IntegerfromAccessTrueFalse = 0
    If Feldname = True Then get_IntegerfromAccessTrueFalse = 1

End Function

Public Sub FeldNameAendern()
On Error GoTo err_FeldNameAendern

Dim sForm As Form
Dim sField As Field
Dim lctl As Control
Dim lres As String
Dim iFeldIndex As Integer


DoCmd.OpenForm "Fehlererfassung - Register", acDesign

Set sForm = Forms("Fehlererfassung - Register")

'Set Control = sForm.Controls("txtFirma")
'frm.SetFocus

For Each lctl In sForm.Controls
    iFeldIndex = 0
    If Not lctl Is Nothing Then
        lres = ""
        lres = FieldDataInRecordset(sForm.RecordSource, lctl.ControlSource)
        If lres <> "" Then
            If iFeldIndex = 0 Then
                lctl.Name = lres
            Else
next_Index:
                lctl.Name = lres & "#" & iFeldIndex
            End If
        lctl.ControlSource = ""
        lctl.ControlTipText = lres
        End If
    End If
err_next:
Next
DoCmd.Close acForm, sForm.Name, acSaveYes

Exit Sub

err_FeldNameAendern:
    If Err = 438 Then
        Err.Clear
        Resume Next
    ElseIf Err = 2104 Then
        Err.Clear
        iFeldIndex = iFeldIndex + 1
        Resume next_Index
    Else

    End If

End Sub
Public Function FieldDataInRecordset(frs As String, fFieldNameRS As String) As String
'Gibt bei positiver Rückmeldung den SourceTabel-Name zurück, ansonsten Leer = ""
On Error GoTo err_FieldDataInRecordset

Dim i As Integer
Dim lrs As QueryDef
Dim lfn As String
Dim ltab As String

    If fFieldNameRS <> "" Then
'        Debug.Print fFieldNameRS
    End If
    Set lrs = CurrentDb.QueryDefs(frs)
    lfn = lrs.Fields(fFieldNameRS).SourceField
    ltab = lrs.Fields(fFieldNameRS).SourceTable

    FieldDataInRecordset = ltab & "#" & lfn

    Exit Function

err_FieldDataInRecordset:
    If Err = 3 Then
        FieldDataInRecordset = ""
    End If

End Function

'##########################################################################################################################
'##############   Funktionen zum Laden der Formular-Controls nach dem Start (open_Form)
'##########################################################################################################################
Public Function IsControlFormCtl(fCtl As Object) As Boolean
'Funktion ermittelt ob das Control einem Formular und nicht einer Registerkarte zugeordnet ist.
'Rückgabewert: True = Control gehört Formular
'Rückgabewert: False = Control liegt auf einem Registerblatt (ControlType = 124)

    If TypeOf fCtl.Parent Is Form Then
        IsControlFormCtl = True
    Else
        If fCtl.Parent.ControlType = 124 Then 'Page Type
            IsControlFormCtl = False
        Else
            IsControlFormCtl = IsControlFormCtl(fCtl.Parent)
        End If
    End If

End Function

Public Function LoadControlsFormAfterOpen(ByRef fForm As Form, ByVal fRecordsource As String, fCriterium As String) As Boolean

    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset
    Dim lctrl As Control
    Dim lSource() As String

    Set ldb = openAccDBConnection("openDB")
    Set lrs = ldb.OpenRecordset("Select * from [" & fForm.getRecordSource() & " ] Where " & fCriterium & ";", dbOpenSnapshot)
    For Each lctrl In fForm.Controls
'        If lctrl.Name = "ReklMemo#WAWE-AB" Or lctrl.Name = "Mastertab#per Telefon#1" Then
'            Debug.Print lctrl.Name
'        End If
        If IsControlFormCtl(lctrl) Then
            lSource = Split(lctrl.Name, "#")
            If UBound(lSource) > 0 Then
'                Debug.Print Nz(DLookup("[" & lSource(1) & "]", "[" & fForm.RecordSource & "]", fCriterium))
'                lctrl = Nz(DLookup("[" & lSource(1) & "]", "[" & fForm.RecordSource & "]", fCriterium))
                lctrl = lrs(lSource(1))
            End If
        End If

    Next
    Set lrs = Nothing
    LoadControlsFormAfterOpen = True

End Function
Public Function LoadControlsPageAfterOpen(ByRef fForm As Form, ByVal fRecordsource As String, ByRef fCtrlList As Object, fCriterium As String) As Boolean

    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset
    Dim lctrl As Control
    Dim lSource() As String

    Set ldb = openAccDBConnection("openDB")
    Set lrs = ldb.OpenRecordset("Select * from [" & fRecordsource & "] Where " & fCriterium & ";", dbOpenSnapshot)
    For Each lctrl In fCtrlList
'        If lctrl.Name = "ReklMemo#WAWE-AB" Or lctrl.Name = "Mastertab#per Telefon#1" Then
'            Debug.Print lctrl.Name
'        End If
        If Not IsControlFormCtl(lctrl) Then
            lSource = Split(lctrl.Name, "#")
            If UBound(lSource) > 0 Then
'                Debug.Print Nz(DLookup("[" & lSource(1) & "]", "[" & fForm.RecordSource & "]", fCriterium))
'                lctrl = Nz(DLookup("[" & lSource(1) & "]", "[" & fForm.RecordSource & "]", fCriterium))
                lctrl = lrs(lSource(1))
            End If
        End If

    Next
    Set lrs = Nothing
    LoadControlsPageAfterOpen = True

End Function

Public Function AddControlListPage(ByRef fForm As Form, ByRef fPage As Page, ByRef fTargetList As ListBox) As Boolean
'Function füllt die Control-List in einer List-Box, die in einem für die Funktion undefinierten Registerkarter auf einem Formular angezeigt werden.
'z.B.: frm_ControlsSteuerung; frm_CommandButtonsSteuerung
'Quelle der List-Box: fPage.Controls
'Target Controls: fTargetList


    Dim ldb As DAO.Database
    Dim lrs As Recordset
    Dim li_ctl As Control
    Dim li_index As Integer
    Dim AddControlList As String

    AddControlList = True
    li_index = 0
    For Each li_ctl In fPage.Controls
        If li_ctl.ControlType = 111 Or li_ctl.ControlType = 109 Or li_ctl.ControlType = 106 Or li_ctl.ControlType = 107 Or li_ctl.ControlType = 104 Then
            'if Control in Tabelle eingetragen
            Set ldb = openAccDBConnection("openDB")
            Set lrs = ldb.OpenRecordset("Select * from [tblControls] where [ControlName] = '" & li_ctl.Name & "' And [FormName] = '" & fForm.Name & "' ;", dbOpenSnapshot)
            If lrs.RecordCount = 0 Then
                ldb.Execute ("Insert into tblControls (FormName, ControlName) values ('" & fForm.Name & "', '" & li_ctl.Name & "');")
                Set lrs = Nothing
            End If
            fTargetList.AddItem Item:=li_ctl.Name, Index:=li_index
            li_index = li_index + 1
        End If
    Next

End Function

Public Function AddControlListForm(ByRef fForm As Form, ByRef fTargetList As ListBox) As Boolean
'Function füllt die Control-List in einer List-Box, die in einem für die Funktion undefinierten Formular, aber nicht auf einer Registerkarte angezeigt werden.
'z.B.: frm_ControlsSteuerung; frm_CommandButtonsSteuerung
'Quelle der List-Box: fForm.Controls
'Target Controls: fTargetList

    Dim ldb As DAO.Database
    Dim lrs As Recordset
    Dim li_ctl As Control
    Dim li_index As Integer
    Dim AddControlList As String

    AddControlList = True
    li_index = 0
    For Each li_ctl In fForm.Controls
        'Abfrage ob sich das Control auf dem Formular und nicht auf einer Registerkarte befindet
        If IsControlFormCtl(li_ctl) Then
            If li_ctl.ControlType = 111 Or li_ctl.ControlType = 109 Or li_ctl.ControlType = 106 Or li_ctl.ControlType = 107 Or li_ctl.ControlType = 104 Then
                'if Control in Tabelle eingetragen
                Set ldb = openAccDBConnection("openDB")
                Set lrs = ldb.OpenRecordset("Select * from [tblControls] where [ControlName] = '" & li_ctl.Name & "' And [FormName] = '" & fForm.Name & "' ;", dbOpenSnapshot)
                If lrs.RecordCount = 0 Then
                    ldb.Execute ("Insert into tblControls (FormName, ControlName) values ('" & fForm.Name & "', '" & li_ctl.Name & "');")
                    Set lrs = Nothing
                End If
                fTargetList.AddItem Item:=li_ctl.Name, Index:=li_index
                li_index = li_index + 1
            End If
        End If
    Next

End Function

Public Function getCtrlDataset(ByRef fForm As Form, fCtrl As Control, Optional fAfterUpdate As Boolean = False) As DAO.Recordset

    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset
    Dim lrsAb As DAO.Recordset
    Dim ls_sub As Variant

    Set ldb = openAccDBConnection("openDB")
    'Eigenschaften des aktuellen übergeben Control ermitteln
    'Unabhängig der Abhängigkeit muss die Änderung direkt auf dem Control in die FieldList aufgenommen werden
    If Not fAfterUpdate Then addControlsList fCtrl.Name, fCtrl.ControlType, fCtrl, fForm
    Set lrs = ldb.OpenRecordset("Select * from tblControls where [FormName] = '" & fForm.Name & "' and [ControlName] = '" & fCtrl.Name & "';", dbOpenSnapshot)
    If lrs.RecordCount Then
        If fAfterUpdate And lrs("Update") Then
            'Set value des aktuell übergeben Control (fCtrl) in Form-Dictionary
            addControlsList fCtrl.Name, fCtrl.ControlType, fCtrl, fForm
            Call Forms(fForm.Name).frmFRControlEventAfterUpdate(fCtrl.Name)
            Exit Function
        End If
        If lrs("Requery") Then
            fForm.Controls(fCtrl.Name).Requery
        End If

        'Ermitteln aller Abhängigkeiten
        Set lrsAb = ldb.OpenRecordset("Select * from tblAbControls where CtrID = " & lrs("CtrID") & ";", dbOpenSnapshot)
        If lrsAb.RecordCount Then
            lrsAb.MoveFirst
            While Not lrsAb.EOF
                Set lrs = ldb.OpenRecordset("Select * from tblControls where [CtrID] = " & lrsAb("AbCtrID") & ";")
                Set lrs = getCtrlDataset(Forms(lrs("FormName")), Forms(lrs("FormName")).Controls(lrs("ControlName")), True)
                lrsAb.MoveNext
            Wend
        End If
    Else
        Set getCtrlDataset = Nothing
    End If

    Set lrs = Nothing

End Function

Public Function CmbEinblenden(strTag As String, fForm As Form)
    Dim ctl As Control
        For Each ctl In fForm.Controls
            If (ctl.ControlType = 104) And (ctl.Tag = strTag) Then
              ctl.Visible = True
            End If
'            If ctl.Tag = stMarker Then ctl.Visible = False
        Next ctl
End Function

Public Function getKurzelZuRekla() As Boolean
'Funktion zur Bestimmung welcher Mitarbeiter(Kürzel) die aktuelle Reklamation gesperrt hat

    Dim Agemeldet As String

        Agemeldet = Nz(DLookup("[Kürzel]", "[Mastertab-DSBerechtigung]", "[ID] = " & get_p_ReklNr & " and [Zustand] = 'Gesperrt' And [Datum] = Date()"))

        If Agemeldet = get_pKuerzel() Then
            getKurzelZuRekla = True
        Else
            getKurzelZuRekla = False
        End If

End Function
Public Function getCmbView(ByRef fForm As Form, fView As String) As DAO.Recordset

    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset
    Dim fCtrl As Control

    Set ldb = openAccDBConnection("openDB")
    'Eigenschaften des aktuellen übergeben Control ermitteln
    Set lrs = ldb.OpenRecordset("Select * from tblControls where [FormName] = '" & fForm.Name & "' and [View] = '" & fView & "';", dbOpenSnapshot)
    If lrs.RecordCount Then
        lrs.MoveFirst
        While Not lrs.EOF
            On Error Resume Next
            fForm.Controls(lrs("ControlName")).Visible = True
            lrs.MoveNext
        Wend
    End If

    Set lrs = Nothing

End Function

Public Function setCmbViewFalse(ByRef fForm As Form) As DAO.Recordset

    Dim ldb As DAO.Database
    Dim lrs As DAO.Recordset
    Dim fCtrl As Control

    Set ldb = openAccDBConnection("openDB")
    'Eigenschaften des aktuellen übergeben Control ermitteln
    Set lrs = ldb.OpenRecordset("Select * from tblControls where [FormName] = '" & fForm.Name & "' and [View] <> '';", dbOpenSnapshot)
    If lrs.RecordCount Then
        lrs.MoveFirst
        While Not lrs.EOF
            On Error Resume Next
            fForm.Controls(lrs("ControlName")).Visible = False
            lrs.MoveNext
        Wend
    End If

    Set lrs = Nothing

End Function

Public Sub VorgangsOrdner(Reklnr As Long, ReklnrKnd As String, Auftragsdatum As Date, OpenPfad As Boolean)      '20190820-ROS
'Vorgangsordner wird angelegt wenn nicht schon vorhanden und geöffnet wenn OpenPfad = True
'Code neu programmiert, da er dreimal vorhanden war, in  unterschiedlicher Ausführung

        Dim stAppName As String
        Dim sPfadDatei As String
        Dim Jahr As String

        Jahr = format(Auftragsdatum, "yyyy")

        If IsNull(ReklnrKnd) Or (ReklnrKnd) = "" Then
            sPfadDatei = getReklVerzeichnis(Reklnr, Jahr)
        Else
            sPfadDatei = getReklVerzeichnis(Reklnr, Jahr, ReklnrKnd)
        End If
        'Prüfung ob das Reklamationsverzeichnis mit einer Meldungsnummer vorhanden ist
        If sPfadDatei = "" Then
            'Nicht vorhanden
            'Prüfen ob das Reklamationsverzeichnis  nur mit der Reklnr. vorhanden ist
            If getReklVerzeichnis(Reklnr, Jahr) = "" Then
                'Nicht Vorhanden
                'Anlegen des Verzeichnisses in Abhängigkeit einer vorh. Meldungsnummer
                If IsNull(ReklnrKnd) Or ReklnrKnd = "" Then
                    sPfadDatei = setReklVerzeichnis(Reklnr, Jahr)
                Else
                    sPfadDatei = setReklVerzeichnis(Reklnr, Jahr, ReklnrKnd)
                End If
            Else
                If Not IsNull(ReklnrKnd) Then
                    sPfadDatei = setReklVerzeichnis(Reklnr, Jahr, ReklnrKnd)
                End If
            End If
        End If

        If OpenPfad = True Then
            stAppName = "C:\Windows\explorer.exe /e," & sPfadDatei
            'Öffnen des Verzeichnisses
            Call shell(stAppName, 1)
        End If

    End Sub
Public Function SonderzeichenEntfernen(s As String) As String    '20190902-ROS

    'Entfernt alle Sonderzeichen aus einem String
    Dim objRegEx As Object
    Set objRegEx = CreateObject("VBscript.regexp")

    objRegEx.IgnoreCase = True
    objRegEx.Global = True
    objRegEx.MultiLine = True
    objRegEx.Pattern = "[^0-9a-zA-Z-_ ]"
'    s = Replace(s, ":", "_")
    SonderzeichenEntfernen = objRegEx.Replace(s, "")
    Set objRegEx = Nothing

End Function

Public Function getPointerPos(ByRef fX As Long, ByRef fY As Long) As Long

    GetCursorPos MousePosition
    fX = MousePosition.X
    fY = MousePosition.Y

End Function
Public Function getValueFromInputBox(ByVal fCaption As String, ByVal fText As String, Optional fgetKuerzel As Boolean = False) As String
On Error GoTo err_getValueFromInputBox
'Öffnet das Formular frmInputbox an den Koordinaten wo die Maus betätigt wurde. Rückgabe des Werts


    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long
    Dim strWert As String
    Dim varKoord As Variant
    Dim strDaten As String
    Dim varKoordInputbox As Variant
    Dim lsOpenArgs As String


    If fgetKuerzel = True Then
        lsOpenArgs = "Form:frmInputBox|Text:" & fText & "|Kuerzel:True|Caption:" & fCaption '& "|"
    Else
        lsOpenArgs = "Form:frmInputBox|Text:" & fText & "|Kuerzel:False|Caption:" & fCaption '& "|"
    End If


    psgetValueFromInputBox = ""
    DoCmd.OpenForm "frmInputBox", _
                    WindowMode:=acDialog, _
                    OpenArgs:=lsOpenArgs
    getValueFromInputBox = psgetValueFromInputBox

'    Forms("frmInputBox").Caption = fCaption
'    Forms("frmInputBox").Controls("txtBoxtext") = fText
'
'    strDaten = Nz(DLookup("Koordinaten", "Tab_Koordinaten", "[Formname] = '" & Forms("frmInputBox").Name & "'"))
'    varKoordInputbox = Split(strDaten, "/")
'
'
'    lgPosition = getPointerPos(XPos, YPos)
'
'    'Positionierung beim öffnen von Access bei ersten mal Kürzel Eingabe.
'    If fgetKuerzel = True Then
'        strWert = Nz(DLookup("[Koordinaten]", "TAB_Koordinaten", "[Formname] = 'TB:autostart'"))
'        varKoord = Split(strWert, "/")
'        FormOpenKoorMousPointer Forms("frmInputBox"), varKoord(0) + 300, varKoord(1) + 160, varKoordInputbox(2), varKoordInputbox(3), False
'    Else
'        FormOpenKoorMousPointer Forms("frmInputBox"), XPos - 190, YPos - 80, varKoordInputbox(2), varKoordInputbox(3), False '380, 160
'    End If
'
'
'    bolStop = True
'    Forms("frmInputBox").SetFocus
'    While bolStop
'     DoEvents
'    Wend
'
'
'    If IsNull(Forms("frmInputBox").Controls("txtBoxValue")) Or Forms("frmInputBox").Controls("txtBoxValue") = "" Then
'        getValueFromInputBox = ""
'    Else
'        getValueFromInputBox = Forms("frmInputBox").Controls("txtBoxValue")
'    End If
'
'    closeWinHandle Forms("frmInputBox")

    Exit Function

err_getValueFromInputBox:

    If Err = 2501 Then
        Resume Next
    End If

End Function

Public Function AufwandsKosten(ByVal fReklID As Long, Optional Ausgaben As Boolean = True) As Double

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT Sum(tblKostenaufstellung.Kosten) AS SummevonKosten " & _
            "FROM tblKostenaufstellung LEFT JOIN tblKostenTyp ON tblKostenaufstellung.KostentypID = tblKostenTyp.KostenTypID " & _
            "WHERE (((tblKostenaufstellung.ReklID)=" & fReklID & ") AND ((tblKostenTyp.Ausgaben)=True));"

    If Ausgaben = False Then
        sql = Replace(sql, "True", "False")
    End If

    Set ldb = openAccDBConnection("openDB")
'    Set rs = ldb.OpenRecordset("SELECT Sum(tblKostenaufstellung.Kosten) AS SummevonKosten " & _
'                            "FROM tblKostenaufstellung LEFT JOIN tblKostenTyp ON tblKostenaufstellung.KostentypID = tblKostenTyp.KostenTypID " & _
'                            "WHERE (((tblKostenaufstellung.ReklID)=" & fReklID & ") AND ((tblKostenTyp.Ausgaben)=True));", dbOpenSnapshot)
    Set rs = ldb.OpenRecordset(sql, dbOpenSnapshot)

    If Not IsNull(rs("SummevonKosten")) Then
        AufwandsKosten = rs("SummevonKosten")
    Else
        AufwandsKosten = 0
    End If
    Set rs = Nothing

End Function

Public Sub TagEintrag(ByVal fReklID As Long, ByVal fKundeID As String)  '20220218-ROS
'Informationen die wichtig für den Kunden sind, werden automatisch in den Tag geladen
'hier ist es der Versandtermin und die AB Terminierung

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs2 As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [RSP_ReklTag] where [Tag] Like  '" & fKundeID & "*';")

    If rs.RecordCount Then
        rs.MoveLast
        rs.MoveFirst
        Do While Not rs.EOF
'            Debug.Print rs("TagID")
            Set rs2 = ldb.OpenRecordset("Select * from [ReklIDtoTagID] where [ReklID] = " & fReklID & " and [TagID] = " & rs("TagID") & " ;")
            If rs2.RecordCount Then
                rs.MoveNext
            Else
                ldb.Execute ("Insert into ReklIDtoTagID ([ReklID], [TagID]) values (" & fReklID & ", " & rs("TagID") & ");")
                rs.MoveNext
            End If
        Loop
    End If

    Forms("Fehlererfassung - Register").Controls("ListReklTag").Requery

    Set rs = Nothing
    Set rs2 = Nothing

End Sub

Public Sub Sanduhr(ByVal fOnOff As String)
'Sanduhr Ein und Aus schalten ON = An Off = Aus

    If fOnOff = "ON" Then DoCmd.Hourglass True
    If fOnOff = "OFF" Then DoCmd.Hourglass False

End Sub

Public Function ZahlAusString(ByVal strGiven As String) As String
On Error Resume Next
    Dim intLength As Integer
    Dim lngValue As Long

        For intLength = 1 To Len(strGiven)
          If IsNumeric(MID(strGiven, intLength, 1)) Then
            lngValue = lngValue & MID(strGiven, intLength, 1)
          End If
        Next intLength
        ZahlAusString = lngValue

End Function


Public Sub InfoProduktion(Optional ByVal fArt As String, Optional ByVal fVar As String, Optional ByVal fFA As String)
'Abrufen von Produktionsdaten bzw. Infos zu den gefertigten FA´s letzten 4 Jahre sortiert nach FA-Nr.

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql_1 As String
    Dim ls_sql_2 As String
    Dim ls_sql As String
    Dim ls_Datum As Date



    ls_sql_1 = "SELECT aint_nr as FA ,tm_artikel.ast_nummer as Art,tm_artikel_var.avar_nummer as Var ,tm_artikel.ast_bezeichnung as Bezeichnung , " & _
                "tt_auftrag_intern.aint_nr as Haupt_FA, tt_auftrag_intern.aint_text3 as FA_Linie, um_produktionsauftrag.NAME4 as KZ,um_produktionsauftrag.DATUM4 as Datum, " & _
                "tt_auftrag_intern_text.aint_memo2 as Memo_Fertigung,tt_auftrag_intern.aint_bedarf_menge as Menge, " & _
                "um_produktionsauftrag.KindOfShipment as Sonder_Info, tt_auftrag_intern.aint_bez2 as Bez2,tt_auftrag_intern_info.atxt_bez4 as Aenderungs_Stand, " & _
                "tt_auftrag_intern.aint_stl_datum as STL_Stand, tt_auftrag_intern_info.ast_zeichnungsnr as Zeich_Nr, tt_auftrag_intern_info.ast_zeichnr_index as Zeich_Index " & _
                "From tt_auftrag_intern " & _
                "INNER JOIN tt_auftrag_intern_info ON tt_auftrag_intern.aint_id = tt_auftrag_intern_info.aint_id " & _
                "INNER JOIN um_produktionsauftrag ON tt_auftrag_intern.aint_id = um_produktionsauftrag.aint_id " & _
                "INNER JOIN tt_auftrag_intern_text ON tt_auftrag_intern.aint_id = tt_auftrag_intern_text.aint_id " & _
                "INNER JOIN tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
                "INNER JOIN tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
                "WHERE tm_artikel.ast_nummer = '###Art###' and tm_artikel_var.avar_nummer = '###Var###' and tt_auftrag_intern.vgst_id = 500 and um_produktionsauftrag.DATUM4 > '###Datum###' " & _
                "ORDER BY FA DESC;"


    ls_sql_2 = "SELECT aint_nr as FA ,tm_artikel.ast_nummer as Art,tm_artikel_var.avar_nummer as Var ,tm_artikel.ast_bezeichnung as Bezeichnung , " & _
                "tt_auftrag_intern.aint_nr as Haupt_FA, tt_auftrag_intern.aint_text3 as FA_Linie, um_produktionsauftrag.NAME4 as KZ,um_produktionsauftrag.DATUM4 as Datum, " & _
                "tt_auftrag_intern_text.aint_memo2 as Memo_Fertigung, tt_auftrag_intern.aint_bedarf_menge as Menge, " & _
                "um_produktionsauftrag.KindOfShipment as Sonder_Info, tt_auftrag_intern.aint_bez2 as Bez2,tt_auftrag_intern_info.atxt_bez4 as Aenderungs_Stand, " & _
                "tt_auftrag_intern.aint_stl_datum as STL_Stand, tt_auftrag_intern_info.ast_zeichnungsnr as Zeich_Nr, tt_auftrag_intern_info.ast_zeichnr_index as Zeich_Index " & _
                "From tt_auftrag_intern " & _
                "INNER JOIN tt_auftrag_intern_info ON tt_auftrag_intern.aint_id = tt_auftrag_intern_info.aint_id " & _
                "INNER JOIN um_produktionsauftrag ON tt_auftrag_intern.aint_id = um_produktionsauftrag.aint_id " & _
                "INNER JOIN tt_auftrag_intern_text ON tt_auftrag_intern.aint_id = tt_auftrag_intern_text.aint_id " & _
                "INNER JOIN tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
                "INNER JOIN tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
                "WHERE FA Like '%###FA###%';"


    If fFA = "" Then
        'Abgeschlossenen Aufträge der letzten 4 Jahre wenn keine FA-NR eingetragen ist
        ls_Datum = DateAdd("yyyy", -4, Date)

        ls_sql_1 = Replace(ls_sql_1, "###Art###", fArt)
        ls_sql_1 = Replace(ls_sql_1, "###Var###", fVar)
        ls_sql_1 = Replace(ls_sql_1, "###Datum###", format(ls_Datum, "yyyy-mm-dd"))
        ls_sql = ls_sql_1
    Else
        'Ausgabe Daten vom FA wenn FA-Nr eingetragen ist
        ls_sql_2 = Replace(ls_sql_2, "###FA###", fFA)
        ls_sql = ls_sql_2
    End If

    Set ldb = openAccDBConnection("openDB")
    'Löschen der Datensätze in Temp-Tabelle
    ldb.Execute ("delete * from tblTempInfoArt")
    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)
    DoCmd.Hourglass True
    'Temp-Tabelle erstellen mit den Daten aus AVISTA
    If rsAvista.RecordCount Then
        Do While Not rsAvista.EOF
            On Error Resume Next
            ldb.Execute ("Insert into tblTempInfoArt ([FA], [Art],[Var], [Bezeichnung], [Haupt_FA], [FA_Linie],[KZ],[Datum],[Memo_Fertigung], [Sonder_Info], " & _
            "[Bez2],[Aenderungs_Stand],[STL_Stand],[Zeich_Nr], [Zeich_Index],[Menge])" & _
            "values('" & rsAvista("FA") & "', '" & rsAvista("Art") & "', '" & rsAvista("Var") & "' , '" & rsAvista("Bezeichnung") & "', '" & rsAvista("Haupt_FA") & "', " & _
            "'" & rsAvista("FA_Linie") & "', '" & rsAvista("KZ") & "', '" & rsAvista("Datum") & "','" & rsAvista("Memo_Fertigung") & "','" & rsAvista("Sonder_Info") & "', " & _
            "'" & rsAvista("Bez2") & "', '" & rsAvista("Aenderungs_Stand") & "', '" & rsAvista("STL_Stand") & "','" & rsAvista("Zeich_Nr") & "','" & rsAvista("Zeich_Index") & "','" & rsAvista("Menge") & "');")
            rsAvista.MoveNext
        Loop
    Else
        MsgBox getTC_FCFromCode("TC4203"), vbInformation, getTC_FCFromCode("TC4204")
        Set rsAvista = Nothing
        Set adbprov = Nothing
        DoCmd.Hourglass False
        Exit Sub
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing

    If IsLoaded("frmInfoArt") Then
        Forms("frmInfoArt").Requery
    Else
        DoCmd.OpenForm ("frmInfoArt")
    End If

    DoCmd.Hourglass False

End Sub

Public Sub InfosZuBG(ByVal fTyp As String, Optional ByVal fArtNr As String, Optional ByVal fFA As String, Optional ByVal fLI As String, Optional ByVal fBstNr As String)
'Ermittlung der Gewichte von ausgelieferten Baugruppen zur Kontrolle auf Vollständigkeit

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim str_LIFaNr As String
    Dim ls_Datum As Date
    Dim l_ls_sql As Boolean
    Dim ls_sql_0 As String

    l_ls_sql = False

    ls_sql = "SELECT aint_nr,tm_artikel.ast_nummer,tm_artikel_var.avar_nummer,tm_artikel.ast_bezeichnung, um_lqpp_fa_cable_weight.cw_weight, " & _
            "tt_auftrag_intern.aint_menge_gelief, um_lqpp_fa_cable_weight.cw_time_date, " & _
            "tt_lftbkopf.lf_nummer As Lieferschein, tt_lftbkopf.lf_best_ref as Bestellnummer, tt_auftrag_intern_text.aint_memo2 as Memo " & _
            "From tt_auftrag_intern " & _
            "INNER JOIN tt_auftrag_intern_text ON tt_auftrag_intern.aint_id = tt_auftrag_intern_text.aint_id " & _
            "INNER JOIN tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
            "INNER JOIN tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
            "INNER JOIN um_lqpp_fa_cable_weight ON tt_auftrag_intern.aint_id = um_lqpp_fa_cable_weight.aint_id " & _
            "INNER JOIN ut_lfpos_fa_referenz ON tt_auftrag_intern.aint_id = ut_lfpos_fa_referenz.aint_id " & _
            "INNER JOIN tt_lftbkopf ON tt_lftbkopf.lf_id = ut_lfpos_fa_referenz.lf_id " & _
            "WHERE tm_artikel.ast_nummer = '" & fArtNr & "';"

    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

    If rsAvista.RecordCount Then
        ls_sql = "SELECT aint_nr,tm_artikel.ast_nummer,tm_artikel_var.avar_nummer,tm_artikel.ast_bezeichnung, um_lqpp_fa_cable_weight.cw_weight, " & _
            "tt_auftrag_intern.aint_menge_gelief, um_lqpp_fa_cable_weight.cw_time_date, " & _
            "tt_lftbkopf.lf_nummer As Lieferschein, tt_lftbkopf.lf_best_ref as Bestellnummer, tt_auftrag_intern_text.aint_memo2 as Memo " & _
            "From tt_auftrag_intern " & _
            "INNER JOIN tt_auftrag_intern_text ON tt_auftrag_intern.aint_id = tt_auftrag_intern_text.aint_id " & _
            "INNER JOIN tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
            "INNER JOIN tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
            "INNER JOIN um_lqpp_fa_cable_weight ON tt_auftrag_intern.aint_id = um_lqpp_fa_cable_weight.aint_id " & _
            "INNER JOIN ut_lfpos_fa_referenz ON tt_auftrag_intern.aint_id = ut_lfpos_fa_referenz.aint_id " & _
            "INNER JOIN tt_lftbkopf ON tt_lftbkopf.lf_id = ut_lfpos_fa_referenz.lf_id " & _
            "###Suchen### = '###NR###';"
    Else
        ls_sql = "SELECT aint_nr,tm_artikel.ast_nummer,tm_artikel_var.avar_nummer,tm_artikel.ast_bezeichnung," & _
                "tt_lftbkopf.lf_nummer As Lieferschein, tt_lftbkopf.lf_best_ref as Bestellnummer, tt_auftrag_intern_text.aint_memo2 as Memo " & _
                "From tt_auftrag_intern " & _
                "INNER JOIN tt_auftrag_intern_text ON tt_auftrag_intern.aint_id = tt_auftrag_intern_text.aint_id " & _
                "INNER JOIN tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
                "INNER JOIN tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
                "INNER JOIN ut_lfpos_fa_referenz ON tt_auftrag_intern.aint_id = ut_lfpos_fa_referenz.aint_id " & _
                "INNER JOIN tt_lftbkopf ON tt_lftbkopf.lf_id = ut_lfpos_fa_referenz.lf_id " & _
                "###Suchen### = '###NR###';"
        l_ls_sql = True
    End If


    If fTyp = "FA" Then
        str_LIFaNr = fFA
        ls_sql = Replace(ls_sql, "###Suchen###", "WHERE tt_auftrag_intern.aint_nr")
    End If

    If fTyp = "LI" Then
        str_LIFaNr = fLI
        ls_sql = Replace(ls_sql, "###Suchen###", "WHERE Lieferschein")
    End If

    If fTyp = "BstNr" Then
        str_LIFaNr = fBstNr
        ls_sql = Replace(ls_sql, "###Suchen###", "WHERE Bestellnummer")
    End If

    If fTyp = "ArtNr" Then
        str_LIFaNr = fArtNr
        ls_sql = Replace(ls_sql, "###Suchen###", "WHERE tm_artikel.ast_nummer")
        ls_sql = Replace(ls_sql, ";", " ORDER BY tt_auftrag_intern.aint_nr DESC;")
    End If

    ls_sql = Replace(ls_sql, "###NR###", str_LIFaNr)

    Set ldb = openAccDBConnection("openDB")
    'Löschen der Datensätze in Temp-Tabelle
    ldb.Execute ("delete * from tblTempInfoBG")
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)
    DoCmd.Hourglass True
    If rsAvista.RecordCount = 0 Then
        If fTyp = "FA" Then
            ls_sql_0 = "SELECT aint_nr ,ast_nummer, avar_nummer, tm_artikel.ast_bezeichnung,  um_lqpp_fa_cable_weight.cw_weight, " & _
                    "um_lqpp_fa_cable_weight.fa_menge_st , um_lqpp_fa_cable_weight.cw_time_date " & _
                    "From DBA.um_lqpp_fa_cable_weight " & _
                    "INNER JOIN DBA.tt_auftrag_intern ON tt_auftrag_intern.aint_id = um_lqpp_fa_cable_weight.aint_id " & _
                    "INNER JOIN DBA.tm_artikel ON tm_artikel.ast_id = tt_auftrag_intern.ast_id " & _
                    "INNER JOIN DBA.tm_artikel_var ON tm_artikel_var.ast_id = tt_auftrag_intern.ast_id AND tm_artikel_var.avar_id = tt_auftrag_intern.avar_id " & _
                    "WHERE aint_nr = '" & str_LIFaNr & "' ORDER BY aint_nr DESC;"
            Set rsAvista = adbprov.getRecordsetAvista(ls_sql_0)
        End If
        If rsAvista.RecordCount = 0 Then
            MsgBox getTC_FCFromCode("TC4202"), vbInformation, getTC_FCFromCode("IC4201")
            Set rsAvista = Nothing
            Set adbprov = Nothing
            DoCmd.Hourglass False
            Exit Sub
        Else
            Do While Not rsAvista.EOF
                ls_Datum = format(rsAvista("cw_time_date"), "dd-mm-yyyy")
                ldb.Execute ("Insert into tblTempInfoBG ([FA], [ArtNr],[Variante], [Bezeichnung],[Kg],[Menge],[Datum],[Lieferschein],[Bestellnummer], [InfoFertigung])" & _
                "values('" & rsAvista("aint_nr") & "', '" & rsAvista("ast_nummer") & "', '" & rsAvista("avar_nummer") & "' , '" & rsAvista("ast_bezeichnung") & "', " & _
                "'" & rsAvista("cw_weight") & "', '" & rsAvista("fa_menge_st") & "' , '" & ls_Datum & "', " & _
                "'','','');")
                rsAvista.MoveNext
            Loop
        End If
    Else
         Do While Not rsAvista.EOF
            If l_ls_sql = False Then
                ls_Datum = format(rsAvista("cw_time_date"), "dd-mm-yyyy")
                ldb.Execute ("Insert into tblTempInfoBG ([FA], [ArtNr],[Variante], [Bezeichnung],[Kg],[Menge],[Datum],[Lieferschein],[Bestellnummer], [InfoFertigung])" & _
                "values('" & rsAvista("aint_nr") & "', '" & rsAvista("ast_nummer") & "', '" & rsAvista("avar_nummer") & "' , '" & rsAvista("ast_bezeichnung") & "', " & _
                "'" & rsAvista("cw_weight") & "', '" & rsAvista("aint_menge_gelief") & "' , '" & ls_Datum & "', " & _
                "'" & rsAvista("Lieferschein") & "','" & rsAvista("Bestellnummer") & "','" & rsAvista("Memo") & "');")
            Else
                ldb.Execute ("Insert into tblTempInfoBG ([FA], [ArtNr],[Variante], [Bezeichnung],[Lieferschein],[Bestellnummer], [InfoFertigung])" & _
                "values('" & rsAvista("aint_nr") & "', '" & rsAvista("ast_nummer") & "', '" & rsAvista("avar_nummer") & "' , '" & rsAvista("ast_bezeichnung") & "', " & _
                "'" & rsAvista("Lieferschein") & "','" & rsAvista("Bestellnummer") & "','" & rsAvista("Memo") & "');")
            End If
            rsAvista.MoveNext
        Loop

    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing
    DoCmd.OpenForm "frmInfoBG"
    Forms("frmInfoBG").Requery
    DoCmd.Hourglass False

End Sub

Public Sub Abfrage_ExcelExport(ByVal fAbfrage As String, ByVal fTbName As String)
'Übergabe des Abfragenamens und Übergabe des Tabellennamen für die Ausgabe der Abfrage nach Excel

    Dim appExcel As Excel.Application
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, fAbfrage, format(Now(), "yyyy/mm/dd_hh.mm") & "_" & fTbName & ".xlsx", True, _
             "Export_Übersicht"

    ShellExecute 0, "open", format(Now(), "yyyy/mm/dd_hh.mm") & "_" & fTbName & ".xlsx", "", "", 1

    Set wks = Nothing
    Set wkb = Nothing

End Sub

Public Sub Stueli_Auslesen(ByRef fsql As String, ByVal fNr As String, Optional ByVal fStueliTyp As String)
'20240311-ROS: Ausgabe der Stücklisten aus AVISTA. Wenn Parameter "FA" gesetzt ist, wir die Stückliste ausgegeben mit der der FA gefertigt wurde,
'ansonsten wir die aktuelle Stückliste ausgegeben.

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim lArt_Fa_Nr As String

    lArt_Fa_Nr = fNr

    If fStueliTyp = "" Then
        Dim ls_Datum As String

        ls_Datum = format(Date, "yyyy-mm-dd")
        fsql = Replace(fsql, "###Datum###", ls_Datum)
    End If

    If IsLoaded("frmStueli2") Then
        DoCmd.Close acForm, "frmStueli2"
    End If



    Set ldb = openAccDBConnection("openDB")
    'Löschen der Datensätze in Temp-Tabelle
    ldb.Execute ("delete * from tblTempStueli")
    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(fsql)

    If rsAvista.RecordCount Then
    'Eingabe der Stüli in Temp-Tabelle
            Do While Not rsAvista.EOF
            ldb.Execute ("Insert into tblTempStueli ([Pos], [Menge],[Artikel],[Var], [Kd_Nr], [Bez], [Typ],[Label],[Hersteller])" & _
            "values('" & rsAvista("Pos") & "', '" & rsAvista("Menge") & "', '" & rsAvista("Artikel") & "' , '" & rsAvista("Var") & "', '" & rsAvista("Kd_Nr") & "', " & _
            "'" & rsAvista("Bez") & "', '" & rsAvista("Typ") & "',  '" & rsAvista("Label") & "', '" & rsAvista("Hersteller") & "');")
            rsAvista.MoveNext
        Loop
        ' Zwei Unterschiedliche Formulare, damit man die Stüli's vergleichen kann
        If fStueliTyp = "FA" Then
            DoCmd.OpenForm ("frmStueliFA")
            Forms!frmStueliFA!txtFaNr = lArt_Fa_Nr
        Else
            If IsLoaded("frmStueli") Then
                DoCmd.OpenForm ("frmStueli2")
                Forms!frmStueli2!txtArtNr = lArt_Fa_Nr
            Else
                DoCmd.OpenForm ("frmStueli")
                Forms!frmStueli!txtArtNr = lArt_Fa_Nr
            End If
        End If
    Else
        MsgBox getTC_FCFromCode("TC4220"), vbInformation, getTC_FCFromCode("TC4221")
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing

End Sub
Public Sub aktuelle_Stueli_abrufen(Optional ByVal fArtNr As String)
'20240311-ROS: Abrufen der aktuellen Stückliste

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim ls_strArtNr As String
    Dim ls_sql_KD As String
    Dim ls_Datum As String

    ls_Datum = format(Date, "yyyy-mm-dd")
    ls_sql = cStueli
    ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", 73)
    ls_sql = Replace(ls_sql, "###Datum###", ls_Datum)

    If fArtNr = "" Then
        ls_strArtNr = InputBox(getTC_FCFromCode("IC4227"), getTC_FCFromCode("IC4228"))
    Else
        ls_strArtNr = fArtNr
    End If

    If ls_strArtNr = "" Then
        Exit Sub
    Else
        ls_sql = Replace(ls_sql, "###Art###", ls_strArtNr)
        Set adbprov = New clsAvistaProvider
        Set rsAvista = New ADODB.Recordset
        Set rsAvista = adbprov.getRecordsetAvista(ls_sql)
        If rsAvista.RecordCount < 1 Then
        'wenn nichts gefunden, wird Suche unter Artikelnummer Kunden wird gestartet
            ls_sql_KD = "SELECT FIRST ast_nummer as ArtNrLQ " & _
                "From DBA.tm_artikel " & _
                "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikel.ast_id = tm_artikelvar_kd.ast_id " & _
                "WHERE tm_artikelvar_kd.avkd_artnr_kd = '###KD###';"
            ls_sql_KD = Replace(ls_sql_KD, "###KD###", ls_strArtNr)
            Set rsAvista = adbprov.getRecordsetAvista(ls_sql_KD)
            ' Wenn was gefunden wird, wir die LQ-Artikelnummer ausgegeben und im String ersetzt
            If rsAvista.RecordCount Then
                ls_sql = cStueli
                ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", 73)
                ls_sql = Replace(ls_sql, "###Art###", rsAvista("ArtNrLQ"))
                ls_sql = Replace(ls_sql, "###Datum###", ls_Datum)
            End If
        End If
        Call Stueli_Auslesen(ls_sql, ls_strArtNr)
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing
End Sub

Public Function AVISTA_ID_KDnr(ByVal fReklID As String) As Long
'Ermittlung der Avista-Kunden-ID anhand der Avista-Kunden-Nr aus der Reklamation

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT Mastertab.ID, tblFirma.KndNr_Avista, tblZuordnung.Zuordnung_ID " & _
                "FROM ((Mastertab LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]) " & _
                "LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID) " & _
                "LEFT JOIN tblZuordnung ON tblFirma.Zurodnung_ID = tblZuordnung.Zuordnung_ID " & _
                "where Mastertab.ID = " & fReklID & " ;")

    If rs("Zuordnung_ID") = 1 Then
        AVISTA_ID_KDnr = 73
    Else
        ls_sql = "SELECT adk_id as KD_ID FROM tm_kunde where adk_debitor_nr = '" & rs("KndNr_Avista") & "';"

        Set adbprov = New clsAvistaProvider
        Set rsAvista = New ADODB.Recordset
        Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

        AVISTA_ID_KDnr = rsAvista("KD_ID")
    End If

    Set rs = Nothing
    Set rsAvista = Nothing
    Set adbprov = Nothing

End Function

Public Function BestNr_Avista(ByVal fABNR As String)
' Ermittlung der Bestellnummer aus AVIST Übergabewert ist die AB-NR

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String

    ls_sql = "SELECT tt_auftrag.af_best_ref as BestNr " & _
            "FROM DBA.tt_abtbkopf INNER JOIN DBA.tt_auftrag ON tt_abtbkopf.af_id = tt_auftrag.af_id " & _
            "WHERE tt_abtbkopf.ma_id = 2 AND tt_abtbkopf.ab_nummer = '#BestNr#';"

    ls_sql = Replace(ls_sql, "#BestNr#", fABNR)


    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

    If rsAvista.RecordCount Then
        BestNr_Avista = rsAvista("BestNr")
    Else

    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing

End Function
Public Sub ErstelleTempTabellen_ADODBSource(ByVal rstSource As ADODB.Recordset, ByVal strTarget As String)

    Dim tabelleVorhanden As Boolean
    Dim ldb As DAO.Database
    Dim tdfNewDef As DAO.TableDef
    Dim fldNewField As DAO.Field
    Dim rstTarget As DAO.Recordset
    Dim i As Integer, j As Integer
    Dim s As String

    If IsLoaded("frmTemp") Then
        DoCmd.Close acForm, "frmTemp"
    End If

    On Error GoTo Fehler_TempTabelleErstellung

    Set ldb = openAccDBConnection("openDB")

    ' Schauen ob Tabelle Vorhanden
    For i = 0 To ldb.TableDefs.Count - 1
        If ldb.TableDefs(i).Name = strTarget Then
            DoCmd.DeleteObject acTable, strTarget
'            tabelleVorhanden = True
        End If
    Next i

    'Vorhanden.dann Daten löschen
    If tabelleVorhanden = True Then
        DoCmd.DeleteObject acTable, strTarget
        tabelleVorhanden = False
    End If

    ' Anlegen Tabelle falls nicht vorhanden
'    If tabelleVorhanden = False Then
         Set tdfNewDef = ldb.CreateTableDef(strTarget)
           For i = 0 To rstSource.Fields.Count - 1
              Dim testType As Integer
              If rstSource.Fields(i).Type = 20 Then
                testType = adSingle
              Else
                testType = adVariant
              End If
              Set fldNewField = tdfNewDef.CreateField(rstSource.Fields(i).Name, testType, rstSource.Fields(i).ActualSize) 'ActualSize
              tdfNewDef.Fields.Append fldNewField

           Next i
        ldb.TableDefs.Append tdfNewDef
'    End If

    ' Füllen Tabelle
    Set rstTarget = ldb.OpenRecordset(strTarget, dbOpenDynaset)

    If Not rstSource.EOF Then rstSource.MoveFirst

    Do While Not rstSource.EOF
        rstTarget.AddNew
        For j = 0 To rstSource.Fields.Count - 1
            If rstSource.Fields(j).Value = "" Then
                rstTarget.Fields(j).Value = " "
            Else
                '20260306-ROS: Absatz in Text entfernt zur besseren Ansicht in Temp-Formular
                s = rstSource.Fields(j).Value
                s = Replace(s, vbCrLf, " ")
                s = Trim(s)
                rstTarget.Fields(j).Value = s
'                rstTarget.Fields(j).Value = rstSource.Fields(j).Value
            End If
        Next j
        rstTarget.Update
        rstSource.MoveNext
    Loop



    Exit Sub

Fehler_TempTabelleErstellung:

       Select Case Err
          Case 3010
             MsgBox "The table " & strTarget & " already exists."
          Case Else
             Resume Next
       End Select

       Exit Sub

End Sub

Public Sub TVN_Abrufen(Optional ByVal fArtNr As String, Optional ByVal fArtVar As String)
'Programmcode zum Abrufen der Teileverwendung für Artikel

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_ArtNr As String
    Dim sql As String
    Dim sql_1 As String

    If fArtNr = "" Then
'        ls_ArtNr = InputBox("Bitte Artikelnummer eingeben", "Artikelnummer")
        ls_ArtNr = InputBox(getTC_FCFromCode("IC4227"), getTC_FCFromCode("IC4265"))
        If ls_ArtNr Like "*-*" Then
            fArtVar = Right(ls_ArtNr, InStr(ls_ArtNr, "-") - 3)
            ls_ArtNr = Left(ls_ArtNr, InStr(ls_ArtNr, "-") - 1)
        End If
        sql = Replace(cs_TVN, "###ArtNr###", ls_ArtNr)
    Else
        ls_ArtNr = fArtNr
        sql = Replace(cs_TVN, "###ArtNr###", ls_ArtNr)
    End If
    If ls_ArtNr = "" Then Exit Sub

    If fArtVar <> "" Then
        sql = Replace(cs_TVN_1, "###ArtNr###", ls_ArtNr)
        sql = Replace(sql, "###ArtVar###", fArtVar)
    End If


    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(sql)

    If rsAvista.RecordCount Then
        If fArtVar <> "" Then
            sql_1 = Replace(cs_TVN_2, "###ast_id###", rsAvista("ast_id"))
            sql_1 = Replace(sql_1, "###avar_id###", rsAvista("avar_id"))
        Else
            sql_1 = Replace(cs_TVN_3, "###ast_id###", rsAvista("ast_id"))
        End If

        Set rsAvista = Nothing
        Set rsAvista = adbprov.getRecordsetAvista(sql_1)

        If rsAvista.RecordCount Then
            'Erstellen Temp-Tabelle
            Call ErstelleTempTabellen_ADODBSource(rsAvista, "TBL_TEMP")
            'Erstellen Temp-Formular
            Call FormularEinrichten("TEMP", "TBL_TEMP")
        Else
            MsgBox "Keine TVN für Artikel - " & ls_ArtNr & " - vorhanden", , "Abfrage TVN"
        End If
    Else
        MsgBox "Keine TVN für Artikel - " & ls_ArtNr & " - vorhanden", , "Abfrage TVN"
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing

End Sub

Public Sub Info_SNR(Optional ByVal fSNR As String)
'Abrufen von Informationen zu Seriennummern aus AVISTA

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql_1 As String
    Dim ls_sql_2 As String
    Dim ls_SNR As String

    If fSNR = "" Then
        ls_SNR = InputBox("Bitte Seriennummer eingeben", "Seriennummer")
    Else
        ls_SNR = fSNR
    End If

    If ls_SNR = "" Then Exit Sub

    ls_sql_1 = Replace(cs_SNR_1, "###SNR###", ls_SNR)
    ls_sql_2 = Replace(cs_SNR_2, "###SNR###", ls_SNR)


    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql_2)

    If rsAvista.RecordCount = 0 Then
        Set rsAvista = adbprov.getRecordsetAvista(ls_sql_1)
        If rsAvista.RecordCount Then
            'Erstellen Temp-Tabelle
            Call ErstelleTempTabellen_ADODBSource(rsAvista, "TBL_TEMP")
            'Erstellen Temp-Formular
            Call FormularEinrichten("TEMP", "TBL_TEMP")
        Else
'            MsgBox ("Kein Datensatz vorhanden")
            MsgBox getTC_FCFromCode("TC4240"), vbInformation, getTC_FCFromCode("TC4241")
        End If
    Else
        'Erstellen Temp-Tabelle
        Call ErstelleTempTabellen_ADODBSource(rsAvista, "TBL_TEMP")
        'Erstellen Temp-Formular
        Call FormularEinrichten("TEMP", "TBL_TEMP")
        DoCmd.OpenForm "frmTEMP", acNormal
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing


End Sub

Public Sub GD1_Art_Daten_Abrufen(Optional ByVal fArtNr As String)

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim Response As Integer
    Dim sql As String

    sql = "SELECT tm_artikel.ast_matchcode as Matchcode, ast_nummer as Art_Nr, tm_artikel_var.avar_nummer as Variante, tm_artikel.ast_bezeichnung as Bezeichnung, tm_artikel.ast_bezeichnung2 as Bezeichnung_2, " & _
        "tm_artikel_var.avar_hersteller as Hersteller, tm_artikel_var.avar_herst_nummer as Hersteller_Nr,tx_artikelstat_bez.artstat_bez as Status, tm_artikel.ast_merkmal_1, tm_artikel.ast_merkmal_2, " & _
        "tm_artikel.ast_merkmal_3 , tm_artikel.ast_merkmal_4, tm_artikel.ast_merkmal_5, tm_artikel.ast_merkmal_6, tm_artikel.ast_merkmal_7 " & _
        "FROM (((DBA.tm_artikel LEFT OUTER JOIN DBA.xm_artikel ON xm_artikel.ast_id = tm_artikel.ast_id INNER JOIN DBA.tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id) " & _
        "LEFT OUTER JOIN DBA.tm_artikel_var_ean ON tm_artikel_var.ast_id = tm_artikel_var_ean.ast_id AND tm_artikel_var.avar_id = tm_artikel_var_ean.avar_id) " & _
        "LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
        "LEFT OUTER JOIN DBA.tm_partei ON tm_partei.pa_id = tm_artikelvar_kd.adk_id " & _
        "LEFT OUTER JOIN DBA.tm_sml_daten ON tm_artikel.ast_id =tm_sml_daten.ast_id AND sprach_id = 1) " & _
        "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
        "LEFT OUTER JOIN DBA.tm_stueckliste ON tm_artikel.ast_id = tm_stueckliste.ast_id " & _
        "LEFT OUTER JOIN DBA.tx_artikelstat_bez ON tm_artikel.artstat_id = tx_artikelstat_bez.artstat_id and tx_artikelstat_bez.sprach_id = 1 " & _
        "WHERE ast_nummer <> '-' AND tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_nummer = '###ArtNr###' AND tm_artikel.ast_id <> 7688 and tm_artikel_var.avar_nummer = 'xxxxx'"

    If fArtNr = "" Then
        fArtNr = InputBox(getTC_FCFromCode("IC4227"), getTC_FCFromCode("TC4230"))
        If fArtNr = "" Then
            MsgBox getTC_FCFromCode("TC4231"), vbInformation, getTC_FCFromCode("TC4230")
            Exit Sub
        End If
    End If

    sql = Replace(sql, "###ArtNr###", fArtNr)

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Fehlerausgabe unterdrücken
    Response = acDataErrAdded
    'Recordset-Instanz
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(sql)

    Select Case rsAvista.RecordCount
        Case 0
            sql = Replace(sql, "and tm_artikel_var.avar_nummer = 'xxxxx'", "")
            Set rsAvista = adbprov.getRecordsetAvista(sql)
    End Select

    If rsAvista.RecordCount Then
        Call ErstelleTempTabellen_ADODBSource(rsAvista, "TBL_TEMP")
        Call FormularEinrichten("TEMP", "TBL_TEMP")
        DoCmd.OpenForm "frmTEMP", acNormal
    Else
        MsgBox getTC_FCFromCode("TC4229"), vbInformation, getTC_FCFromCode("TC4230")
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing


End Sub
Public Sub cmbPruefprotokoll_Oeffnen(ByVal fFA_NR As String)

    Dim DeineDatei As String
    Dim PfadDateiName  As String
    Dim FA_NR As String
    Dim Konfektion As String
    Dim Pfad As String
    Dim Jahr As String
    Dim strDir As String
    Dim objFSO As Object
    Dim objDir As Object
    Dim DatOrd As Variant
    Dim Datei As Object
    Dim PDF As Boolean
    Dim i As Integer

    PDF = False

    fFA_NR = LTrim(fFA_NR)
    fFA_NR = RTrim(fFA_NR)

    DeineDatei = fFA_NR & "*" & ".pdf"

    Jahr = "20" & MID(fFA_NR, 4, 2)

    Pfad = "F:\Fertigung\Prüfprotokolle_Kabelkonfektion\" & Jahr & "\"

    i = 1
    DoCmd.Hourglass True
    For i = 0 To 2
        If Dir(Pfad & DeineDatei) <> "" Then
            DeineDatei = Dir(Pfad & DeineDatei)
            PfadDateiName = Pfad & DeineDatei
            ShellExecute 0, "Open", PfadDateiName, "", "", 1
            PDF = True
            Exit For
        Else
            Set objFSO = CreateObject("scripting.filesystemobject")
            strDir = Pfad
            Set objDir = objFSO.GetFolder(strDir)

            For Each DatOrd In objDir.SubFolders        'Unterordner durchsuchen
                If Dir(DatOrd & "\" & DeineDatei) <> "" Then
                    DeineDatei = Dir(DatOrd & "\" & DeineDatei)
                    PfadDateiName = DatOrd & "\" & DeineDatei
                    ShellExecute 0, "Open", PfadDateiName, "", "", 1
                    PDF = True
                    Exit For
                End If
            Next
        End If
        'wenn nichts gefunden, im Vorgangsordner ein Jahr später suchen
        If PDF = False Then
            i = i + 1
            Jahr = Jahr + 1
            Pfad = "F:\Fertigung\Prüfprotokolle_Kabelkonfektion\" & Jahr & "\"
        Else
            Exit For
        End If
    Next i
    DoCmd.Hourglass False
    If PDF = False Then MsgBox getTC_FCFromCode("TC4252"), vbInformation, getTC_FCFromCode("IC4249")

    Set objDir = Nothing
    Set objFSO = Nothing

End Sub

Public Sub TerminStellungnahmeKundeLQ()

    Dim rs As New ADODB.Recordset
    rs.Open "qryTerminStellungnahmeKundeLQ", CurrentProject.Connection, _
            adOpenKeyset, adLockOptimistic

    If rs.RecordCount Then
        'Erstellen Temp-Tabelle
        Call ErstelleTempTabellen_ADODBSource(rs, "TBL_TEMP")
        'Erstellen Temp-Formular
        Call FormularEinrichten("TEMP", "TBL_TEMP")
    Else
        MsgBox getTC_FCFromCode("TC4240"), vbInformation, getTC_FCFromCode("TC4208")
    End If

    Set rs = Nothing

End Sub
Public Sub ErstelleTempTabellen_DAOSource(ByVal rstSource As DAO.Recordset, ByVal strTarget As String)

    Dim tabelleVorhanden As Boolean
    Dim ldb As DAO.Database
    Dim tdfNewDef As DAO.TableDef
    Dim fldNewField As DAO.Field
    Dim rstTarget As DAO.Recordset
    Dim i As Integer, j As Integer

'    tabelleVorhanden = False

    On Error GoTo Fehler_TempTabelleErstellung

    Set ldb = openAccDBConnection("openDB")

    If IsLoaded("frmTemp") Then
        DoCmd.Close acForm, "frmTemp"
    End If

    ' Schauen ob Tabelle Vorhanden
    For i = 0 To ldb.TableDefs.Count - 1
        If ldb.TableDefs(i).Name = strTarget Then
            DoCmd.DeleteObject acTable, strTarget
'            tabelleVorhanden = True
        End If
    Next i

    'Vorhanden.dann Daten löschen
'    If tabelleVorhanden = True Then
'        DoCmd.DeleteObject acTable, strTarget
'        tabelleVorhanden = False
'    End If

    ' Anlegen Tabelle falls nicht vorhanden
'    If tabelleVorhanden = False Then
         Set tdfNewDef = ldb.CreateTableDef(strTarget)
           For i = 0 To rstSource.Fields.Count - 1
              Dim testType As Integer
              If rstSource.Fields(i).Type = 20 Then
                testType = adSingle
              Else
                testType = adVariant
              End If
              Set fldNewField = tdfNewDef.CreateField(rstSource.Fields(i).Name, testType, rstSource.Fields(i).Size) 'ActualSize
              tdfNewDef.Fields.Append fldNewField

           Next i
        ldb.TableDefs.Append tdfNewDef
'    End If

    ' Füllen Tabelle
    Set rstTarget = ldb.OpenRecordset(strTarget, dbOpenDynaset)

    If Not rstSource.EOF Then rstSource.MoveFirst

    Do While Not rstSource.EOF
        rstTarget.AddNew
        For j = 0 To rstSource.Fields.Count - 1
            If rstSource.Fields(j).Value = "" Then
                rstTarget.Fields(j).Value = " "
            Else
                rstTarget.Fields(j).Value = rstSource.Fields(j).Value
            End If
        Next j
        rstTarget.Update
        rstSource.MoveNext
    Loop



    Exit Sub

Fehler_TempTabelleErstellung:

       Select Case Err
          Case 3010
             MsgBox "The table " & strTarget & " already exists."
          Case Else
             Resume Next
       End Select

       Exit Sub

End Sub

Public Sub Rekla_Zu_BGArt_Bez(Optional fBG_Bez As String)
' Reklamationen nach Artikelbezeichnung in Baugruppen und Artikel suchen

    Dim lsqlBG As String
    Dim lsqlArt As String
    Dim lBezeichnung As String
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    If fBG_Bez = "" Then
        lBezeichnung = "*" & InputBox(getTC_FCFromCode("IC4253"), getTC_FCFromCode("IC4254")) & "*"
    Else
        lBezeichnung = "*" & fBG_Bez & "*"
    End If

    lsqlBG = "SELECT Mastertab.ID, [Baugruppen-avista_out].[LQ Art-Nr] As BG, [Baugruppen-avista_out].[Bezeichnung 1]" & _
            "FROM Mastertab LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]" & _
            "GROUP BY Mastertab.ID, [Baugruppen-avista_out].[LQ Art-Nr], [Baugruppen-avista_out].[Bezeichnung 1]" & _
            "Having ((([Baugruppen-avista_out].[Bezeichnung 1]) like '" & lBezeichnung & "'));"

    lsqlArt = "SELECT Mastertab.ID, [Artikel-all_avista_out].[LQ Art-Nr] As ArtNr, [Artikel-all_avista_out].[Bezeichnung 1]" & _
            "FROM Mastertab LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]" & _
            "GROUP BY Mastertab.ID, [Artikel-all_avista_out].[LQ Art-Nr], [Artikel-all_avista_out].[Bezeichnung 1]" & _
            "HAVING ((([Artikel-all_avista_out].[Bezeichnung 1]) Like '" & lBezeichnung & "'));"

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(lsqlBG)
    If rs.RecordCount Then
        'Erstellen Temp-Tabelle
        Call ErstelleTempTabellen_DAOSource(rs, "TBL_TEMP")
        'Erstellen Temp-Formular
        Call FormularEinrichten("TEMP", "TBL_TEMP")
    Else
        Set rs = ldb.OpenRecordset(lsqlArt)
        If rs.RecordCount Then
            'Erstellen Temp-Tabelle
            Call ErstelleTempTabellen_DAOSource(rs, "TBL_TEMP")
            'Erstellen Temp-Formular
            Call FormularEinrichten("TEMP", "TBL_TEMP")
        Else
            MsgBox getTC_FCFromCode("TC4240"), vbInformation, getTC_FCFromCode("TC3218")
        End If
    End If

    Set rs = Nothing

End Sub

Public Sub UebernahmeP4T()
'Makro zur Überprüfung der Übernahme der Stellungnhame im P4T
'die zu Überprüfende Tabelle muss aus dem P4T herunter geladen werden,
'diese wird dann als Temp-Tabelle im Access angelegt "TempTabP4T"
'die Tabelle "TempTabP4T" wird dann mit der im Access erzeuten Tabelle "tblP4T" verglichen
'Übereinstimmungen werden in der Tabelle "tblP4T" mit 1 gekennzeichnet die anderen mit 0
'danach öffnet sich ein Temp-Formular mit den abgeglichenen Daten aus der Tabelle "tblP4T"

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs_2 As DAO.Recordset
    Dim i As Integer
    Dim strFeld As String
    Dim strExcelFile As String
    Dim strTableName As String
    Dim fld As DAO.Field

    Set ldb = openAccDBConnection("openDB")

    strExcelFile = DateiOeffnen
    If strExcelFile = "" Then Exit Sub

    strTableName = "TempTabP4T"

    DoCmd.DeleteObject acTable, strTableName
    DoCmd.DeleteObject acTable, "tblP4T"
    ldb.Execute "SELECT * INTO tblP4T FROM qryStellungnahme_Klärung_Kunde_Trumpf"

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, strTableName, strExcelFile, True

    Set rs = ldb.OpenRecordset("Select * from " & strTableName & ";", dbOpenSnapshot)
    Set rs_2 = ldb.OpenRecordset("Select * from tblP4T;")

    Do Until rs.EOF
        ' Schleife durch alle Felder (Spalten) des aktuellen Datensatzes
        For Each fld In rs.Fields
            If fld.Value = "Ext. Reklamations Nr." Then
                ' fld.OrdinalPosition gibt die sichtbare Position aus (0-basiert)
                i = fld.OrdinalPosition
                Exit Do ' Beendet die äußere Schleife, wenn der erste Treffer reicht
            End If
        Next fld
        rs.MoveNext
    Loop

    strFeld = rs.Fields(i).Name
'    strFeld = rs.Fields(3).Name

    If rs_2.RecordCount Then
        rs_2.MoveLast
        rs_2.MoveFirst
        Do While Not rs_2.EOF
            If rs_2("Meldungsnummer Kunde") = 6296507 Or rs_2("Meldungsnummer Kunde") = 6312817 Then
                Debug.Print rs_2("Meldungsnummer Kunde")
            End If
            If IsNull(rs_2("Meldungsnummer Kunde")) Then
                rs_2.MoveNext
            Else
                rs_2.Edit
                Set rs = ldb.OpenRecordset("SELECT * From " & strTableName & " WHERE " & strFeld & " Like '*" & rs_2("Meldungsnummer Kunde") & "*';")
                If rs.RecordCount Then
                    rs_2("P4T") = 100
                Else
                    rs_2("P4T") = 0
                End If
                rs_2.Update
                rs_2.MoveNext
            End If
        Loop
    End If

    Set rs = Nothing
    Set rs_2 = Nothing
    Set ldb = Nothing

'    DoCmd.DeleteObject acTable, strTableName

    Call FormularEinrichten("TEMP", "tblP4T")


End Sub

Public Function Buchungstext(ByVal fsDatum As Date, fsLagerID As Integer, fsArtNr As String) As String

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim lDatum As String

    lDatum = format(fsDatum, "yyyy-mm-dd")

    ls_sql = "SELECT ast_nummer + '-' + avar_nummer AS ArtNr, tm_artikel.ast_bezeichnung as Bez, tt_mb_beleg_pos.mbp_menge as St, tt_mb_beleg_kopf.mbk_date as Zugang, mbk_text1 as BuchText " & _
            "FROM ((((DBA.tt_mb_beleg_kopf INNER JOIN DBA.tt_mb_beleg_pos ON tt_mb_beleg_kopf.mbk_id = tt_mb_beleg_pos.mbk_id ) INNER JOIN DBA.tm_artikel  ON tt_mb_beleg_pos.ast_id = tm_artikel.ast_id )" & _
            "INNER JOIN DBA.tm_artikel_var ON tt_mb_beleg_pos.ast_id = tm_artikel_var.ast_id AND tt_mb_beleg_pos.avar_id = tm_artikel_var.avar_id )  LEFT OUTER JOIN DBA.tm_kostenstelle ON tt_mb_beleg_pos.kst_id = tm_kostenstelle.kst_id )" & _
            "LEFT OUTER JOIN DBA.tt_auftrag_intern ON tt_mb_beleg_pos.aint_id = tt_auftrag_intern.aint_id LEFT OUTER JOIN DBA.xm_artikel_ext ON xm_artikel_ext.ast_id = tm_artikel.ast_id " & _
            "WHERE tt_mb_beleg_kopf.mbk_art = 1 and (tt_mb_beleg_pos.mbp_art = 10 or tt_mb_beleg_pos.mbp_art = 1) and tt_mb_beleg_pos.lager_id = ###Lager### and Zugang = '###Datum###' and ArtNr = '" & fsArtNr & "' ;"

    ls_sql = Replace(ls_sql, "###Lager###", fsLagerID)
    ls_sql = Replace(ls_sql, "###Datum###", lDatum)

    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

    If rsAvista.RecordCount Then
        Buchungstext = rsAvista("BuchText")
    Else
        ls_sql = "SELECT ast_nummer + '-' + avar_nummer AS ArtNr, tm_artikel.ast_bezeichnung as Bez, tt_mb_beleg_pos.mbp_menge as St, tt_mb_beleg_kopf.mbk_date as Zugang, mbk_text1 as BuchText " & _
            "FROM ((((DBA.tt_mb_beleg_kopf INNER JOIN DBA.tt_mb_beleg_pos ON tt_mb_beleg_kopf.mbk_id = tt_mb_beleg_pos.mbk_id ) INNER JOIN DBA.tm_artikel  ON tt_mb_beleg_pos.ast_id = tm_artikel.ast_id )" & _
            "INNER JOIN DBA.tm_artikel_var ON tt_mb_beleg_pos.ast_id = tm_artikel_var.ast_id AND tt_mb_beleg_pos.avar_id = tm_artikel_var.avar_id )  LEFT OUTER JOIN DBA.tm_kostenstelle ON tt_mb_beleg_pos.kst_id = tm_kostenstelle.kst_id )" & _
            "LEFT OUTER JOIN DBA.tt_auftrag_intern ON tt_mb_beleg_pos.aint_id = tt_auftrag_intern.aint_id LEFT OUTER JOIN DBA.xm_artikel_ext ON xm_artikel_ext.ast_id = tm_artikel.ast_id " & _
            "WHERE tt_mb_beleg_kopf.mbk_art = 4 and tt_mb_beleg_pos.mbp_art = 41 and tt_mb_beleg_pos.lager_id = ###Lager### and Zugang = '###Datum###' and ArtNr = '" & fsArtNr & "';"

        ls_sql = Replace(ls_sql, "###Lager###", fsLagerID)
        ls_sql = Replace(ls_sql, "###Datum###", lDatum)
        Set adbprov = New clsAvistaProvider
        Set rsAvista = New ADODB.Recordset
        Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

        If rsAvista.RecordCount Then
            Buchungstext = rsAvista("BuchText")
        Else
            Buchungstext = ""
        End If
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing



End Function

Public Function LagerplatzArtikel(fsMatNr As String, Optional fsMatNrVar As String = "00000") As String
On Error GoTo err_LagerplatzArtikel
'Funktion ruft den Lagerplatz in AVISTA auf, übergeben wird die Artikelnummer und die Variante

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim rsAvistaOut As ADODB.Recordset
    Dim sql As String

    Set adbprov = New clsAvistaProvider

    Set rsAvistaOut = adbprov.getLQMatNrBezeichnung(fsMatNr, fsMatNrVar)
    'Rückgabe der Art_ID und der VAR_ID wird für die weitere Verarbeitung im sql benötigt

    If Not rsAvistaOut Is Nothing Then
        sql = "SELECT Top 1  lager_bez as Lager " & _
                "FROM (DBA.tt_bestand_sum INNER JOIN DBA.tm_artikel ON tt_bestand_sum.ast_id = tm_artikel.ast_id) " & _
                "LEFT OUTER JOIN DBA.xm_artikel ON  tt_bestand_sum.ast_id = xm_artikel.ast_id  INNER JOIN DBA.vm_lager ON vm_lager.lager_id = tt_bestand_sum.lager_id " & _
                "WHERE tt_bestand_sum.ast_id = ###AstID### AND tt_bestand_sum.avar_id = ###AstAvar### ORDER BY CASE WHEN tt_bestand_sum.lager_id = 4 THEN 1 ELSE 2 END, lager_bez DESC; "

        sql = Replace(sql, "###AstID###", rsAvistaOut("ast_id"))
        sql = Replace(sql, "###AstAvar###", rsAvistaOut("avar_id"))

        Set rsAvista = adbprov.getRecordsetAvista(sql)

        LagerplatzArtikel = rsAvista("Lager")
    Else
        LagerplatzArtikel = ""
    End If

    Set rsAvistaOut = Nothing
    Set rsAvista = Nothing
    Set adbprov = Nothing

    Exit Function
err_LagerplatzArtikel:

    If Err Then
        Resume Next
    End If

End Function

Public Function AufwandInEuro(ByVal fReklID As Long)
'Errechnet alle Aufwände in Euro mit dem jeweiligen Stundensatz pro Postion

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT AufwandVorgangPositionen.MID, Sum(AufwandVorgangPositionen.Aufwand) AS SummevonAufwand " & _
            "From AufwandVorgangPositionen " & _
            "GROUP BY AufwandVorgangPositionen.MID " & _
            "Having (AufwandVorgangPositionen.MID)= " & fReklID & ";"

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(sql, dbOpenSnapshot)

    If rs.RecordCount Then
        AufwandInEuro = rs("SummevonAufwand")
    Else
        AufwandInEuro = 0
    End If

    Set rs = Nothing

End Function
Public Sub KI_Internet()
'Daten aus Zwischenablage in Internetbrowser kopieren und verarbeiten.


    Dim objIE1 As Object
    Dim ChromeLocation As String
    Dim lUser As String
    Dim lComputername As String
    Dim lHTTPS As String

    lHTTPS = "https://www.google.com/search?q=&rlz=1C1GCEB_enDE986DE986&sourceid=chrome&ie=UTF-8&udm=50&aep=48&cud=0&qsubts=1781097050069&source=chrome.crn.obic"
'    lHTTPS = "https://perplexity.ai/"
'    lHTTPS = "https://openai.com/de-DE"
'    lHTTPS = "https://gemini.google.com/"

    lUser = fOSUserName
    lComputername = Environ("COMPUTERNAME")

    If lComputername = "LQTERMINAL02" Then
        ChromeLocation = "C:\Users\" & lUser & "\AppData\Local\Google\Chrome\Application\chrome.exe"
'        shell (ChromeLocation & " -url " & lHTTPS), vbNormalFocus
'        Sleep 1500             ' 1000 Millisekunden = 1 Sekunde warten
'        SendKeys "^v"          ' Strg + V
'        Sleep 750
'        SendKeys "{ENTER}"     ' optional: Enter drücken
'        SendKeys "{NUMLOCK}"
    Else
        ChromeLocation = "C:\Program Files\Google\Chrome\Application\chrome.exe"
'        shell (ChromeLocation & " -url " & lHTTPS), vbNormalFocus

        '############ MS-Copilot ist nicht auf jedem Rechner verfügbar laut IT #########################
'        Call Copilot_Shortcut_Druecken
'        Sleep 750             ' 1000 Millisekunden = 1 Sekunde warten
'        SendKeys "^v"          ' Strg + V
'        Sleep 750
'        SendKeys "{ENTER}"     ' optional: Enter drücken
'        SendKeys "{NUMLOCK}"
    End If

    shell (ChromeLocation & " -url " & lHTTPS), vbNormalFocus
    Sleep 1500             ' 1000 Millisekunden = 1 Sekunde warten
    SendKeys "^v"          ' Strg + V
    Sleep 750
    SendKeys "{ENTER}"     ' optional: Enter drücken
    SendKeys "{NUMLOCK}"

End Sub

Public Sub Daten_AB_Avista(ByVal fABNR As String)
'20260218-ROS: Abrufen von AB-Daten aus AVISTA

    If fABNR = "" Then Exit Sub

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT ab_nummer as AB_Nr, tt_abtbkopf.adk_nummer as KD_Nr, af_besteller as Besteller, " & _
            "tt_auftrag.af_best_ref as Bestell_Nr, tt_abtbkopf.ab_lfadr_name1 as Lief_Adress_1, " & _
            "tt_abtbkopf.ab_lfadr_name2 as Lief_Adress_2, tt_abtbkopf.ab_lfadr_name3 as Lief_Adress_3, " & _
            "tt_abtbkopf.ab_lfadr_strasse as Strasse, tt_abtbkopf.ab_lfadr_plz as PLZ, " & _
            "tt_abtbkopf.ab_lfadr_ort as Ort, tt_auftrag.af_status as Status_500_Abgeschlossen, " & _
            "tm_user.cmuser_name as Sachbearbeiter_LQ, tt_abtbkopf.ab_netto as Wert, " & _
            "tt_abtbkopf_notiz.ab_text1 as Notiz_1, tt_abtbkopf_notiz.ab_text2 as Notiz_2 " & _
            "From DBA.tt_abtbkopf " & _
            "INNER JOIN DBA.tt_auftrag ON tt_abtbkopf.af_id = tt_auftrag.af_id " & _
            "LEFT OUTER JOIN DBA.tm_user ON tm_user.cmuser_id = tt_auftrag.af_sachbearb " & _
            "LEFT OUTER JOIN DBA.tt_abtbkopf_notiz ON tt_abtbkopf_notiz.ab_id = tt_abtbkopf.ab_id " & _
            "WHERE tt_abtbkopf.ma_id = 2 AND tt_abtbkopf.ab_nummer = '###AB###';"


    sSQL = Replace(sSQL, "###AB###", fABNR)

    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(sSQL)

    If rsAvista.RecordCount Then
        Set rsAvista = adbprov.getRecordsetAvista(sSQL)
        'Erstellen Temp-Tabelle
        Call ErstelleTempTabellen_ADODBSource(rsAvista, "TBL_TEMP")
        'Erstellen Temp-Formular
        Call FormularEinrichten("TEMP", "TBL_TEMP")
        DoCmd.OpenForm "frmTEMP", acNormal
    Else
        'Kein Datensatz vorhanden
         MsgBox getTC_FCFromCode("TC4240"), vbInformation, getTC_FCFromCode("TC4241")
    End If


    Set rsAvista = Nothing
    Set adbprov = Nothing



End Sub


Public Sub setBookmark(ByRef ffrm As Form, ByVal fID As Variant, fFeld As String)

'        llTicketIDSave = Me.TicketID
'        Me.Painting = False
'        Me.Requery
'        Set lrs = Me.RecordsetClone
'        lrs.FindFirst "TicketID = " & llTicketIDSave
'        Me.Bookmark = lrs.Bookmark
'        Me.Painting = True
'        Set lrs = Nothing
        Dim lrs As DAO.Recordset

        Set lrs = ffrm.RecordsetClone
        Select Case lrs(fFeld).Type
            Case 4, 3, 5, 14, 16
                lrs.FindFirst "[" & fFeld & "] = " & fID
            Case Is > 200
                lrs.FindFirst "[" & fFeld & "]'" & fID & "'"
            Case Else
                lrs.FindFirst "[" & fFeld & "]'" & fID & "'"

        End Select
        ffrm.Bookmark = lrs.Bookmark
        Set lrs = Nothing

End Sub

Public Sub SpringeZuProzedur(ByVal ModulName As String, ByVal ProzedurName As String)
'20260803-ROS

    Dim vbeProj As Object
    Dim vbeComp As Object
    Dim codeMod As Object
    Dim startZeile As Long
    Dim anzZeilen As Long

    On Error GoTo FehlerHandler

    ' Zugriff auf das aktive VBA-Projekt
    Set vbeProj = Application.VBE.ActiveVBProject

    ' Modul suchen
    Set vbeComp = vbeProj.VBComponents(ModulName)
    Set codeMod = vbeComp.CodeModule

    ' Entwicklungsumgebung sichtbar machen und Fenster öffnen
    vbeComp.Activate

    ' Startzeile der Prozedur ermitteln (0 = Sub/Function, nicht Deklaration)
    startZeile = codeMod.ProcStartLine(ProzedurName, 0)

    ' Anzahl der Zeilen der Prozedur ermitteln
    anzZeilen = codeMod.ProcCountLines(ProzedurName, 0)

    'VBA-Editor öffnen
    vbeComp.CodeModule.CodePane.Show

    ' Prozedur im Editor anzeigen und den gesamten Block markieren
    ' Parameter: StartZeile, StartSpalte, EndZeile, EndSpalte
    codeMod.CodePane.SetSelection startZeile, 1, startZeile + anzZeilen - 1, 1

    Exit Sub

FehlerHandler:
    MsgBox "Fehler: Modul '" & ModulName & "' oder Prozedur '" & _
           ProzedurName & "' wurde nicht gefunden!", vbCritical, "Fehler"
End Sub


Public Sub ProgrammcodeDurchsuchen(fSuchbegriff As String)
'20260803-ROS

    Dim ldb As DAO.Database
    Dim Modul As Object
    Dim modCode As Object
    Dim startZeile As Long
    Dim startSpalte As Long
    Dim endZeile As Long
    Dim endSpalte As Long
    Dim gefunden As Boolean
    Dim ProzedurName As String
    Dim prozedurTyp As Long
    Dim DatenGefunden As Boolean
    Dim zeilenText As String
    Dim ProzedurArt As String

    If IsLoaded("frmProgrammcode") Then DoCmd.Close acForm, "frmProgrammcode"


    Set ldb = openAccDBConnection("openDB")

    ldb.Execute "DELETE * FROM tblTempProgrammcode;", dbFailOnError

    If Trim(fSuchbegriff) = "" Then Exit Sub

    DoCmd.Hourglass True
    DatenGefunden = False

    ' Schleife durch alle Komponenten (Module, Formulare, Berichte)
    For Each Modul In Application.VBE.ActiveVBProject.VBComponents
        Set modCode = Modul.CodeModule

        ' Suche für das aktuelle Modul initialisieren
        startZeile = 1
        startSpalte = 1
        endZeile = modCode.CountOfLines
        endSpalte = 255

        ' Suche starten
        gefunden = modCode.Find(fSuchbegriff, startZeile, startSpalte, endZeile, endSpalte, True, False, False)

        Do While gefunden
            On Error Resume Next
            ProzedurName = modCode.ProcOfLine(startZeile, prozedurTyp)

            If Err.Number <> 0 Or ProzedurName = "" Then
                ProzedurName = "[Keine Prozedur / Deklarationsbereich]"
            End If

            'Den Text der Startzeile auslesen
            zeilenText = modCode.Lines(modCode.ProcBodyLine(ProzedurName, prozedurTyp), 1)

            'Prüfen, ob "Function" oder "Sub" im Text vorkommt
            If InStr(1, zeilenText, "Function", vbTextCompare) > 0 Then
                ProzedurArt = "Function"
            ElseIf InStr(1, zeilenText, "Sub", vbTextCompare) > 0 Then
                ProzedurArt = "Sub"
            ElseIf InStr(1, zeilenText, "Property", vbTextCompare) > 0 Then
                ProzedurArt = "Property"
            Else
                ProzedurArt = "Unbekannt"
            End If

            'Daten in tblTempProgrammcode eintragen
            ldb.Execute ("Insert into tblTempProgrammcode (Modul, ProzedurName, ProzedurArt) values " & _
            "('" & Modul.Name & "', '" & ProzedurName & "', '" & ProzedurArt & "');")

            ' Weitersuchen ab der nächsten Zeile im selben Modul
            startZeile = startZeile + 1
            startSpalte = 1
            endZeile = modCode.CountOfLines
            endSpalte = 255

            gefunden = modCode.Find(fSuchbegriff, startZeile, startSpalte, endZeile, endSpalte, True, False, False)

            'Zwischenspeichern ob Daten gefunden wurden
            DatenGefunden = True
        Loop
    Next Modul
    DoCmd.Hourglass False

    'Abfragen ob Daten gefunden wurden
    If DatenGefunden = False Then
        MsgBox "Die Suche nach '" & fSuchbegriff & "' war nicht erfolgreich.", vbInformation, "Fertig"
    Else
        DoCmd.OpenForm "frmProgrammcode", , , , , , fSuchbegriff
    End If

    Set modCode = Nothing

End Sub
Public Sub ExportRecordSourceNachExcel(ByVal strRecordSource As String, ByVal strFilter As String, Optional ByVal strCustomFileName As String = "Excel_Export")
'20260825-ROS

    Dim ldb As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim strQueryName As String
    Dim strDateipfad As String
    Dim strDownloadsPath As String

    ' Sicherheitsprüfung: Falls die Datenherkunft leer übergeben wurde
    If Trim(strRecordSource) = "" Then
        MsgBox getTC_FCFromCode("TC4298"), vbExclamation, getTC_FCFromCode("TC0228")
        Exit Sub
    End If

    ' 1. Pfad zum Windows-Download-Ordner auslesen
    strDownloadsPath = Environ("USERPROFILE") & "\Downloads\"

    ' 2. Datenbank-Objekt referenzieren
    Set ldb = openAccDBConnection("openDB")

    ' 3. Name für die temporäre Abfrage und den Zielpfad definieren
    strQueryName = strCustomFileName & "_ExcelExport"
    strDateipfad = strDownloadsPath & strCustomFileName & "_" & format(Now, "yyyymmdd_hhnn") & ".xlsx"

    ' 4. SQL-String basierend auf den übergebenen Parametern zusammenbauen
    If Trim(strFilter) <> "" Then
        strSQL = "SELECT * FROM (" & strRecordSource & ") WHERE " & strFilter
    Else
        strSQL = "SELECT * FROM (" & strRecordSource & ")"
    End If

    ' 5. Falls die temporäre Abfrage existiert, vorher löschen
    On Error Resume Next
    ldb.QueryDefs.Delete strQueryName
    On Error GoTo 0

    ' 6. Temporäre Abfrage mit der übergebenen RecordSource erstellen
    Set qdf = ldb.CreateQueryDef(strQueryName, strSQL)

    ' 7. Export nach Excel ausführen
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strQueryName, strDateipfad, True
    ShellExecute 0, "open", strDateipfad, "", "", 1

    ' 8. Aufräumen: Temporäre Abfrage wieder löschen
    ldb.QueryDefs.Delete strQueryName
    Set qdf = Nothing

    ' Erfolgsmeldung
    'MsgBox "Export erfolgreich abgeschlossen!" & vbCrLf & "Datei gespeichert unter:" & vbCrLf & strDateipfad, vbInformation, "Export beendet"
End Sub


Public Sub PrintForm(ByVal lForm As Form)

    Dim origOrientation As Long

    ' 1. Aktuelle Druckerausrichtung sichern
    origOrientation = Application.Printer.Orientation

    ' 2. Globalen Access-Drucker für diesen Moment auf QUERFORMAT zwingen
    ' (2 = acPRORLandscape)
    Application.Printer.Orientation = 2

    ' 3. Ränder des Formulars anpassen für die Zentrierung (in Twips: 567 Twips = 1 cm)
    With lForm.Printer
        .LeftMargin = 1440   ' ca. 2,5 cm linker Rand
        .RightMargin = 1440  ' ca. 2,5 cm rechter Rand
        .TopMargin = 1134    ' ca. 2,0 cm oberer Rand
    End With

    ' 4. Aktuelles Formular fest im Fokus verankern
    lForm.SetFocus
    DoCmd.SelectObject acForm, lForm.Name, False

    ' 5. ALLE DATENSÄTZE DRUCKEN
    ' acPrintAll zwingt Access, die gesamte Datenbasis auszugeben
    DoCmd.PrintOut acPrintAll

    ' 6. Drucker-Einstellung wieder auf den Ursprungswert zurücksetzen
    Application.Printer.Orientation = origOrientation

End Sub
