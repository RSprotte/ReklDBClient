Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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

Const cs_Gewichte = "SELECT  aint_nr as FA, ast_nummer as ArtNr, avar_nummer as Variante, tm_artikel.ast_bezeichnung as Bezeichnung, um_lqpp_fa_cable_weight.cw_weight as Kg, " & _
                    "um_lqpp_fa_cable_weight.fa_menge_st as Menge ,um_lqpp_fa_cable_weight.cw_time_date as Datum " & _
                    "From DBA.um_lqpp_fa_cable_weight " & _
                    "INNER JOIN DBA.tt_auftrag_intern ON tt_auftrag_intern.aint_id = um_lqpp_fa_cable_weight.aint_id " & _
                    "INNER JOIN DBA.tm_artikel ON tm_artikel.ast_id = tt_auftrag_intern.ast_id " & _
                    "INNER JOIN DBA.tm_artikel_var ON tm_artikel_var.ast_id = tt_auftrag_intern.ast_id AND tm_artikel_var.avar_id = tt_auftrag_intern.avar_id " & _
                    "LEFT JOIN DBA.um_lqpp_artikel_var ON um_lqpp_artikel_var.ast_id = tm_artikel.ast_id AND um_lqpp_artikel_var.avar_id = tm_artikel_var.avar_id " & _
                    "LEFT JOIN DBA.um_lqpp_infosys_data ON um_lqpp_infosys_data.ins_id = 2 AND um_lqpp_infosys_data.ins_key_1 = tm_artikel.ast_id AND um_lqpp_infosys_data.ins_key_2 = 0 " & _
                    "WHERE tm_artikel.ast_nummer = '###Art###' ORDER BY aint_nr DESC"

Const cs_ServiceLager = "SELECT tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_bezeichnung,  " & _
                        "tt_bestand_sum.besum_ist_menge, tm_lager.lager_bez, tt_bestand_sum.besum_zu_dat_last " & _
                        "FROM DBA.tt_bestand_sum INNER JOIN DBA.tm_artikel ON tt_bestand_sum.ast_id = tm_artikel.ast_id " & _
                        "INNER JOIN DBA.tm_artikel_var ON tt_bestand_sum.ast_id = tm_artikel_var.ast_id  AND tt_bestand_sum.avar_id = tm_artikel_var.avar_id " & _
                        "INNER JOIN DBA.tm_lager ON tt_bestand_sum.lager_id =  tm_lager.lager_id INNER JOIN DBA.tm_partei ON tt_bestand_sum.lager_id = tm_partei.pa_id " & _
                        "LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikel_var.ast_id = tm_artikelvar_kd.ast_id AND tm_artikel_var.avar_id = tm_artikelvar_kd.avar_id  " & _
                        "AND tm_artikelvar_kd.avkd_default = 1 AND tm_artikelvar_kd.adk_id = 73 " & _
                        "WHERE tm_lager.ma_id = 2 AND tt_bestand_sum.lager_id = ###Lager### AND tt_bestand_sum.besum_ist_menge >= 1.000000"

Const cs_BuchText = "SELECT tt_mb_beleg_kopf.mbk_text1 " & _
                    "FROM ((((DBA.tt_mb_beleg_kopf INNER JOIN DBA.tt_mb_beleg_pos ON tt_mb_beleg_kopf.mbk_id = tt_mb_beleg_pos.mbk_id )" & _
                    "INNER JOIN DBA.tm_artikel  ON tt_mb_beleg_pos.ast_id = tm_artikel.ast_id )" & _
                    "INNER JOIN DBA.tm_artikel_var ON tt_mb_beleg_pos.ast_id = tm_artikel_var.ast_id AND tt_mb_beleg_pos.avar_id = tm_artikel_var.avar_id )" & _
                    "LEFT OUTER JOIN DBA.tm_kostenstelle ON tt_mb_beleg_pos.kst_id = tm_kostenstelle.kst_id )" & _
                    "WHERE tt_mb_beleg_kopf.mbk_date = '###Datum###' AND tt_mb_beleg_pos.lager_id = ###Lager### AND tm_artikel.ast_nummer = '###Art###' AND tm_artikel_var.avar_nummer = '###Var###' AND tt_mb_beleg_kopf.ma_id = 2"

Const cs_TVN_1 = "SELECT tm_artikel.ast_id, tm_artikel_var.avar_id " & _
                "FROM (DBA.tm_artikel INNER JOIN DBA.tm_artikel_text ON tm_artikel.ast_id = tm_artikel_text.ast_id " & _
                "INNER JOIN DBA.tm_artikel_zuord ON tm_artikel_zuord.ast_id = tm_artikel.ast_id) " & _
                "INNER JOIN DBA.tm_artikel_var ON tm_artikel_var.ast_id = tm_artikel.ast_id " & _
                "WHERE sprach_id = 1 AND atxt_typ=1 AND tm_artikel.ast_nummer = '###ArtNr###';"

Const cs_TVN_2 = "SELECT DISTINCT tm_artikel.ast_nummer as ArtNr, tm_artikel.ast_bezeichnung as Bez, tm_artikel.ast_merkmal_4 as Merkmal " & _
                "FROM ((( DBA.tm_stueckliste INNER JOIN DBA.tm_stueli_posvers ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id) " & _
                "INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id) " & _
                "INNER JOIN DBA.tm_artikel ON tm_stueckliste.ast_id = tm_artikel.ast_id ) " & _
                "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                "WHERE tx_artikelstatus.artstat_archiv = 0 And tm_stueckliste_pos.stlp_member_id = ###ast_id### And tm_stueckliste_pos.stlp_member_sub_id = ###avar_id### " & _
                "ORDER BY ast_nummer;"
                '##############
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


' Hilfsfunktion: Extrahiert den übersetzten Text sauber aus der JSON-Antwort
Private Function ExtrahiereLibreText(ByVal jsonText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim schluessel As String

    schluessel = """translatedText"":"""
    startPos = InStr(jsonText, schluessel)

    If startPos > 0 Then
        startPos = startPos + Len(schluessel)
        endPos = InStr(startPos, jsonText, """")
        If endPos > startPos Then
            ExtrahiereLibreText = MID(jsonText, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    ExtrahiereLibreText = "Fehler: Text konnte nicht extrahiert werden."
End Function

' Hilfsfunktion: Bereinigt Zeilenumbrüche und Anführungszeichen, damit das JSON nicht bricht
Private Function CleanJsonString(ByVal Text As String) As String
    Dim res As String
    res = Replace(Text, "\", "\\")
    res = Replace(res, """", "\""")
    res = Replace(res, vbCrLf, "\n")
    res = Replace(res, vbCr, "\n")
    res = Replace(res, vbLf, "\n")
    CleanJsonString = res
End Function

Public Function GoogleTranslate(Text As String, SourceLang As String, TargetLang As String) As String
    Dim objHTTP As Object
    Dim objJS As Object
    Dim URL As String
    Dim Response As String
    Dim CleanText As String

    ' WICHTIG: Absolut sicheres URL-Encoding für Umlaute, Sonderzeichen und Zeilenumbrüche
    On Error Resume Next
    Set objJS = CreateObject("ScriptControl")
    If objJS Is Nothing Then
        ' Alternative für moderne Windows-Systeme via htmlfile-Objekt
        Set objJS = CreateObject("htmlfile")
        objJS.parentWindow.execScript "function encode(s) { return encodeURIComponent(s); }", "JScript"
        CleanText = objJS.parentWindow.encode(Text)
    Else
        objJS.Language = "JScript"
        CleanText = objJS.Run("encodeURIComponent", Text)
    End If
    On Error GoTo ErrorHandler

    ' Falls das Encoding fehlschlägt, rudimentäre Notfall-Maskierung
    If CleanText = "" Then
        CleanText = Replace(Text, " ", "%20")
    End If

    ' URL zusammensetzen
    URL = "https://googleapis.com" & _
          SourceLang & "&tl=" & TargetLang & "&dt=t&q=" & CleanText

    ' Verbindung aufbauen
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    objHTTP.send

    ' Antwort auswerten
    If objHTTP.Status = 200 Then
        Response = objHTTP.responseText
        If InStr(Response, """") > 0 Then
            GoogleTranslate = Split(Split(Response, """")(1), """")(0)
        End If
    Else
        GoogleTranslate = "Fehler: Status " & objHTTP.Status
    End If

    Set objHTTP = Nothing
    Set objJS = Nothing
    Exit Function

ErrorHandler:
    GoogleTranslate = "Verbindungsfehler: " & Err.Description
    Set objHTTP = Nothing
    Set objJS = Nothing
End Function




Private Sub zum_Testen_Click()
'zum_Testen

'    If Not IsNull(Me.txtOriginal) Then
'        ' Übersetzt von Deutsch ("de") nach Englisch ("en")
'        Me.txtUebersetzung = GoogleTranslate(Me.txtOriginal, "de", "en")
'    End If

    Debug.Print GoogleTranslate("Wie hoch ist der Eifelturm", "de", "en")

End Sub

Private Sub Befehl_Fernglas_Click()
'Test_Fernglas

    ' Aktiviert das aktuelle Formular, um sicherzustellen, dass es das Zielobjekt ist
    DoCmd.SelectObject acForm, Me.Name, False

    ' Druckt das aktive Objekt (das Formular) direkt aus
    DoCmd.PrintOut


End Sub

Private Sub Befehl4_DE_Click() 'Filter löschen

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("qryÜbersichtExtern", dbOpenSnapshot)

    Call ErstelleTempTabellen_DAOSource(rs, "tblÜbersichtExtern")

'    strSQL = "SELECT * INTO " & "tblÜbersichtExtern" & " FROM qryÜbersichtExtern;"
'
'    ' Ausführen des SQL-Befehls
'    ldb.Execute strSQL, dbFailOnError

End Sub


Private Sub cmbGewicht_Click()

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim adbprov2 As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As DAO.Recordset
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim str_Artikel As String
    Dim ls_Datum As Date



    str_Artikel = InputBox(getTC_FCFromCode("IC4200"), getTC_FCFromCode("IC4201"))
'    str_Artikel = "1649070"
    If str_Artikel = "" Then Exit Sub

    ls_sql = Replace(cs_Gewichte, "###Art###", str_Artikel)

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("delete * from tblTempInfoBG")
    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

    If rsAvista.RecordCount Then
        Do While Not rsAvista.EOF
            ls_Datum = format(rsAvista("Datum"), "dd-mm-yyyy")
            ldb.Execute ("Insert into tblTempInfoBG ([FA], [ArtNr],[Variante], [Bezeichnung], [Kg], [Menge],[Datum])" & _
            "values('" & rsAvista("FA") & "', '" & rsAvista("ArtNr") & "', '" & rsAvista("Variante") & "' , " & _
            "'" & rsAvista("Bezeichnung") & "', '" & rsAvista("Kg") & "',  '" & rsAvista("Menge") & "', '" & ls_Datum & "');")
            rsAvista.MoveNext
        Loop
    Else
        MsgBox getTC_FCFromCode("TC4202"), vbInformation, getTC_FCFromCode("IC4201")
        Set rsAvista = Nothing
        Set rs = Nothing
        Set adbprov = Nothing
        Exit Sub
    End If

    DoCmd.OpenForm ("frmInfoBG")

    Set rsAvista = Nothing
    Set rs = Nothing
    Set adbprov = Nothing

End Sub

Private Sub Befehl4_Click()
'On Error GoTo err_nextCollection
'
'    Dim s_coll(2) As String
'    Dim i As Integer
'    Dim ldb As DAO.Database
'    Dim l_ref As Reference
'    Dim s_PDF As String
'    Dim rs As DAO.Recordset
'
'    Set ldb = openAccDBConnection("openDB")
'    Set rs = ldb.OpenRecordset("Select * from tblVerweise;")
'
'    If rs.RecordCount Then
'       rs.MoveFirst
'        Do While Not rs.EOF
'            If IsNull(rs("VerweisName")) And IsNull(rs("VerweisPfadStandard")) Then rs.MoveNext
'            If IsNull(rs("VerweisName")) Then
'                s_PDF = rs("Verweis")
'            Else
'                s_PDF = rs("VerweisName")
'            End If
'
'            If Application.References(s_PDF).IsBroken Then
'
'
'start_AddFromFile:
'
'                s_coll(0) = rs("VerweisPfadStandard")
'                s_coll(1) = rs("VerweisPfadAlternative")
'
'                i = -1
'Next_i:
'                i = i + 1
'                If i < UBound(s_coll) Then
'                    Set l_ref = Application.References.AddFromFile(s_coll(i))
''                    Debug.Print l_ref.Name
'                    rs.Edit
'                    rs("VerweisName") = l_ref.Name
'                    rs.Update
'                    'db.Execute ("Update KonfigurationReklDB set [Value] = '" & l_ref.Name & "' where Konfiguration = 'PDFCreatorName';")
'                End If
'
'                If i >= UBound(s_coll) Then
'                    MsgBox "Verweis von x hat nicht funktioniert!"
'                End If
'            End If
'
'        rs.MoveNext
'        Loop
'        Exit Sub
'    End If
'
'err_nextCollection:
'
'    If Err = 9 Then 'Refrenceobject nicht vorhanden
'        Resume start_AddFromFile
'    ElseIf Err = 29060 Then 'Fehler AddFromFile
'        Resume Next_i
'    End If
'
'    Set rs = Nothing
End Sub

Private Sub Befehl12_Click()
On Error GoTo Err_Befehl12_Click

Dim sDatum1 As Date
Dim sDatum2 As Date
Dim Test As Long
Dim liHour As Integer
Dim liDays As Integer
Dim liMinute As Integer

    sDatum1 = Me.txtDatum1
    sDatum2 = Me.txtDatum2


        Me.txtDiff = CalcArbeitszeitDiffDate(Me.txtDatum1, Me.txtDatum2)
        Me.txtDiff = getHourMinuteByAnzMinute(Me.txtDiff, liDays, liHour, liMinute)

        Me.txtAT = liDays
        Me.txth = liHour
        Me.txtmin = liMinute

Exit_Befehl12_Click:
    Exit Sub

Err_Befehl12_Click:
    MsgBox Err.Description
    Resume Exit_Befehl12_Click

End Sub
Private Sub Test()

    Dim X As String

    X = Forms("tblTicketListe")!cboTicketTyp
    'Debug.Print X

End Sub


Private Sub cmbAbgleichHerststellerNr_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim sql As String
    Dim strArtNr As String
    Dim strVar As String

    DoCmd.Hourglass True
    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Artikel-all_avista_out];")
    rs.MoveLast
    rs.MoveFirst
    Do While Not rs.EOF
        If IsNull(rs("Hersteller-Teilenummer")) Or rs("Hersteller-Teilenummer") = "" Then
            rs.MoveNext
        Else
            strArtNr = rs("LQ Art-Nr")
            If IsNull(rs("Variante")) Then
                strVar = "00000"
            Else
                strVar = rs("Variante")
            End If
'            If strArtNr = "0054603" Then
'                Debug.Print strArtNr
'            End If
            sql = ("SELECT tm_artikel_var.avar_herst_nummer " & _
                "FROM (((DBA.tm_artikel LEFT OUTER JOIN DBA.xm_artikel ON xm_artikel.ast_id = tm_artikel.ast_id INNER JOIN DBA.tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id)" & _
                "LEFT OUTER JOIN DBA.tm_artikel_var_ean ON tm_artikel_var.ast_id = tm_artikel_var_ean.ast_id AND tm_artikel_var.avar_id = tm_artikel_var_ean.avar_id) " & _
                "LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
                "LEFT OUTER JOIN DBA.tm_partei ON tm_partei.pa_id = tm_artikelvar_kd.adk_id LEFT OUTER JOIN DBA.tm_sml_daten ON tm_artikel.ast_id =tm_sml_daten.ast_id AND sprach_id = 1) " & _
                "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id LEFT OUTER JOIN DBA.xm_artikel_var ON tm_artikel_var.ast_id = xm_artikel_var.ast_id " & _
                "AND tm_artikel_var.avar_id = xm_artikel_var.avar_id LEFT OUTER JOIN DBA.tm_artvar_lief ON tm_artikel_var.ast_id = tm_artvar_lief.ast_id AND tm_artikel_var.avar_id = tm_artvar_lief.avar_id " & _
                "AND tm_artvar_lief.avli_prioritaet = 1 LEFT OUTER JOIN DBA.tm_stueckliste ON tm_artikel.ast_id = tm_stueckliste.ast_id  WHERE ast_nummer <> '-' AND tx_artikelstatus.artstat_archiv = 0 " & _
                "AND tm_artikel.ast_nummer = '" & strArtNr & "' AND tm_artikel_var.avar_nummer = '" & strVar & "';")
            Set rsAvista = adbprov.getRecordsetAvista(sql)
            If rsAvista.RecordCount Then
                If rs("Hersteller-Teilenummer") <> rsAvista("avar_herst_nummer") Or IsNull(rs("Hersteller-Teilenummer")) Then
                    rs.Edit
                    rs("Hersteller-Teilenummer") = rsAvista("avar_herst_nummer")
                    rs.Update
                End If
            End If
            rs.MoveNext
            Set rsAvista = Nothing
'        Else
'            rs.MoveNext
        End If
    Loop
    DoCmd.Hourglass False
    Set rs = Nothing
    Set adbprov = Nothing

End Sub



Public Sub cmbBereichToAvista_Click()

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim dbAccess As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim Aktiv As Integer


    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Daten aus Access auslesen
    Set dbAccess = openAccDBConnection("openDB")
    Set rsAccess = dbAccess.OpenRecordset("Select * from LQBereich ;")
        rsAccess.MoveFirst
        Do While Not rsAccess.EOF
            If rsAccess("Aktiv") = True Then
                Aktiv = 1
            Else
                Aktiv = 0
            End If
            adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_bereich ([feb_id],[fee_id],[feb_name],[feb_aktiv])" & _
                                        "Values (" & rsAccess("ID") & "," & rsAccess("ID-LQEinheit") & ", '" & rsAccess("LQBereich") & "'," & Aktiv & ");")
            rsAccess.MoveNext
        Loop

    Set rsAccess = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende

End Sub

Private Sub cmbDefTaetigkeitenToAvista_Click()

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim dbAccess As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim Aktiv As Integer


    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Daten aus Access auslesen
    Set dbAccess = openAccDBConnection("openDB")
    Set rsAccess = dbAccess.OpenRecordset("Select * from ZuordnungDefTaetigkeit ;")
        rsAccess.MoveFirst
        Do While Not rsAccess.EOF
            adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_fehler_info([fef_id],[fet_id],[fef_text]) Values (" & rsAccess("ID") & ", " & rsAccess("TaetigkeitID") & ",'" & rsAccess("DefTaetigkeit") & "');")
            rsAccess.MoveNext
        Loop

    Set rsAccess = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende


End Sub

Public Sub cmbEinheitToAvista_Click()

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim dbAccess As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim Aktiv As Integer


    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Daten aus Access auslesen
    Set dbAccess = openAccDBConnection("openDB")
    Set rsAccess = dbAccess.OpenRecordset("Select * from LQEinheit ;")
        rsAccess.MoveFirst
        Do While Not rsAccess.EOF
            If rsAccess("Aktiv") = True Then
                Aktiv = 1
            Else
                Aktiv = 0
            End If
            adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_einheit ([fee_id],[fee_name],[fee_aktiv]) Values (" & rsAccess("ID") & ", '" & rsAccess("LQEinheit") & "'," & Aktiv & ");")
            rsAccess.MoveNext
        Loop

    Set rsAccess = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende


End Sub


Private Sub cmbServiceLager_Click()
    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim adbprov2 As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As DAO.Recordset
    Dim rsAvista As ADODB.Recordset
    Dim rsAvista2 As ADODB.Recordset
    Dim ls_sql As String
    Dim ls_sql2 As String


        Set ldb = openAccDBConnection("openDB")
        'Temp tabelle leeren
        ldb.Execute ("delete * from tblTempLager ")
        Set rs = ldb.OpenRecordset("Select * from tlbServiceLager ")
        If rs.RecordCount Then
            Do While Not rs.EOF
                ls_sql = ""
                ls_sql = Replace(cs_ServiceLager, "###Lager###", rs("ID"))
                Set adbprov = New clsAvistaProvider
                Set rsAvista = New ADODB.Recordset
                Set rsAvista = adbprov.getRecordsetAvista(ls_sql)
                If rsAvista.RecordCount Then
                    Do While Not rsAvista.EOF
                        ls_sql2 = ""
                        ls_sql2 = Replace(cs_BuchText, "###Lager###", rs("ID"))
                        ls_sql2 = Replace(ls_sql2, "###Datum###", format(rsAvista("besum_zu_dat_last"), "yyyy-mm-dd"))
                        ls_sql2 = Replace(ls_sql2, "###Art###", rsAvista("ast_nummer"))
                        ls_sql2 = Replace(ls_sql2, "###Var###", rsAvista("avar_nummer"))
                        Set adbprov2 = New clsAvistaProvider
                        Set rsAvista2 = New ADODB.Recordset
                        Set rsAvista2 = adbprov.getRecordsetAvista(ls_sql2)
                        ldb.Execute ("Insert into tblTempLager ([ArtNr], [Var],[KdNrTrumpf], [Bez], [St], [Lager],[Eingang],[BuchText])" & _
                        "values('" & rsAvista("ast_nummer") & "', '" & rsAvista("avar_nummer") & "', '" & rsAvista("avkd_artnr_kd") & "' , '" & rsAvista("ast_bezeichnung") & "', '" & rsAvista("besum_ist_menge") & "',  '" & rsAvista("lager_bez") & "', '" & rsAvista("besum_zu_dat_last") & "', '" & rsAvista2("mbk_text1") & "');")
                        rsAvista.MoveNext
                        Set rsAvista2 = Nothing
                    Loop
                End If
                Set rsAvista = Nothing
                rs.MoveNext
            Loop
            ls_sql = ""
        End If


        DoCmd.OpenForm ("frmServiceLager")

        Set db = Nothing
        Set rs = Nothing
        Set adbprov = Nothing

End Sub

Public Sub cmbTaetigkeitToAvista_Click()

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim dbAccess As DAO.Database
    Dim rsAccess As DAO.Recordset


    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Daten aus Access auslesen
    Set dbAccess = openAccDBConnection("openDB")
    Set rsAccess = dbAccess.OpenRecordset("Select * from Taetigkeiten where ([Zuordnung_ID] = 3 or [Zuordnung_ID] = 5) ;")
        rsAccess.MoveFirst
        Do While Not rsAccess.EOF
            adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_taetigkeit ([fet_id],[fet_name]) Values (" & rsAccess("ID") & ", '" & rsAccess("Taetigkeit") & "');")
            rsAccess.MoveNext
        Loop

    Set rsAccess = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende


End Sub



Private Sub Form_Load()

    Dim ctl As Control


    If getCurrLanguage = "German" Then
        For Each ctl In Me.Controls
            If Right(ctl.Name, 3) = "_DE" Then
                ctl.Visible = True
            End If
        Next
    End If
End Sub


Private Function get_DE() As Boolean

    get_DE = True

End Function


Function LokaerKIAbruf(promptText As String) As String
    Dim objHTTP As Object
    Dim JsonPayload As String
    Dim Response As String

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Die URL zeigt auf Ihren lokalen Ollama-Server
    Dim URL As String
    URL = "https://www.google.com/search?q=&rlz=1C1GCEB_enDE986DE986&sourceid=chrome&ie=UTF-8&udm=50&aep=48&cud=0&qsubts=1781097050069&source=chrome.crn.obic"

    ' JSON-Format, angepasst an Ihr lokales Modell
    JsonPayload = "{""model"": ""llama3"", ""prompt"": """ & promptText & """, ""stream"": false}"

    objHTTP.Open "POST", URL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send JsonPayload

    ' Hier müsste die Antwort (ResponseText) ausgelesen und per JSON-Parser zerlegt werden
    Response = objHTTP.responseText

    LokaerKIAbruf = Response
End Function
Function ErsteZahlAusText(ByVal s As String) As Long
    Dim i As Long
    Dim startPos As Long
    Dim tmp As String

    ' Position der ersten Ziffer suchen
    For i = 1 To Len(s)
        If MID$(s, i, 1) Like "#" Then
            startPos = i
            Exit For
        End If
    Next i

    ' Wenn keine Ziffer gefunden wurde, 0 zurückgeben
    If startPos = 0 Then
        ErsteZahlAusText = 0
        Exit Function
    End If

    ' Ab erster Ziffer alle folgenden Ziffern einsammeln
    For i = startPos To Len(s)
        If MID$(s, i, 1) Like "#" Then
            tmp = tmp & MID$(s, i, 1)
        Else
            Exit For
        End If
    Next i

    ErsteZahlAusText = CLng(tmp)
End Function
