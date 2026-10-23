Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Befehl3_Click()

On Error GoTo Err_cmbAktiveHostAnzeigen_Click

        Dim stDocName As String
        Dim stLinkCriteria As String

        stDocName = "frmAktiverHost"
        DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormReadOnly

Exit_cmbAktiveHostAnzeigen_Click:
        Exit Sub

Err_cmbAktiveHostAnzeigen_Click:
        MsgBox Err.Description
        Resume Exit_cmbAktiveHostAnzeigen_Click

End Sub

Private Sub Befehl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Befehl6_Click()

    DoCmd.OpenForm ("frmEmailService")

End Sub

Private Sub Befehl6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cboVorgangsOrdner_Click()
On Error Resume Next

    Dim ReklnrKnd As String
    Dim ReklID As Long
    Dim lDatum As Date

    ReklID = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
    If ReklID = 0 Then
        MsgBox "Keine Vorgangsnummer eingetragen!", , "Vorgangsnummer"
        Exit Sub
    Else
        ReklnrKnd = Nz(DLookup("[Meldungsnummer Kunde]", "Mastertab", "[ID] = " & ReklID & ""))
        lDatum = Nz(DLookup("[Auftragsdatum]", "Mastertab", "[ID] = " & ReklID & ""))
        Call VorgangsOrdner(ReklID, ReklnrKnd, lDatum, True)
    End If

End Sub

Private Sub cboVorgangsOrdner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbABDaten_Click()

    Call Daten_AB_Avista(InputBox(getTC_FCFromCode("IC4266"), getTC_FCFromCode("IC4267")))

End Sub

Private Sub cmbABDaten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbFormular1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbGLS_Click()
'Abrufen der Daten von Paketdienstleister GLS zur Reklamation für die Kostenermittlung

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String


    sql = "SELECT tblKostenaufstellung.ReklID, tblKostenTyp.Anzeige, tblKostenaufstellung.InfoKostenTyp " & _
            "FROM tblKostenaufstellung LEFT JOIN tblKostenTyp ON tblKostenaufstellung.KostentypID = tblKostenTyp.KostenTypID " & _
            "WHERE (((tblKostenaufstellung.KostenTypID)=6 Or (tblKostenaufstellung.KostenTypID)=8 Or (tblKostenaufstellung.KostenTypID)=9) AND ((tblKostenaufstellung.Kosten)=0)) " & _
            "GROUP BY tblKostenaufstellung.ReklID, tblKostenTyp.Anzeige, tblKostenaufstellung.InfoKostenTyp;"

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(sql)
    If rs.RecordCount Then
        'Erstellen Temp-Tabelle
        Call ErstelleTempTabellen_DAOSource(rs, "TBL_TEMP")
        'Erstellen Temp-Formular
        Call FormularEinrichten("TEMP", "TBL_TEMP")
    Else
        MsgBox getTC_FCFromCode("TC4240"), vbInformation, getTC_FCFromCode("TC3218")
    End If

    Set rs = Nothing

End Sub

Private Sub cmbGLS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbInfozuFA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbOffeneAB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbP4TAusgabe_Click()

'    Call Abfrage_ExcelExport("qryStellungnahme_Klärung_Kunde_Trumpf", "Stellungnahme_Klärung_Trumpf")
    Call UebernahmeP4T

End Sub

Private Sub cmbFormular1_Click()

    DoCmd.OpenForm ("Formular1")

End Sub

Private Sub cmbInfozuFA_Click()

    Dim strFaNr As String

    strFaNr = InputBox(getTC_FCFromCode("IC4205"), getTC_FCFromCode("IC4206"))
    If strFaNr = "" Then
        Exit Sub
    Else
        InfoProduktion , , strFaNr
    End If

End Sub

Private Sub cmbP4TAusgabe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbPerplextiy_Click()

    DoCmd.OpenForm ("frmKI")

End Sub

Private Sub cmbPerplextiy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbSNR_Click()

    Call Info_SNR

End Sub

Private Sub cmbSNR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbStueli_Click()

    Call aktuelle_Stueli_abrufen

End Sub

Private Sub cmbStueli_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbTVN_Click()

    Call TVN_Abrufen

End Sub

Private Sub cmbTVN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbUpdate_Click()
    Dim ldb As DAO.Database
    Dim rsRekla As DAO.Recordset
    Dim rsGU As DAO.Recordset
    Dim rsRep As DAO.Recordset
    Dim Zeit As Double


    Set ldb = openAccDBConnection("openDB")

'    Set rs = SELECT Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit From [AufwandVorgangPositionen] WHERE ([AufwandVorgangPositionen.Datum] = Date() AND [AufwandVorgangPositionen.Kürzel]= '" & get_pKuerzel() & "');", dbOpenSnapshot)
    Set rsRekla = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.Datum, Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit " & _
                    "FROM (Mastertab LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID " & _
                    "WHERE (((AufwandVorgangPositionen.Kürzel) = '" & get_pKuerzel() & "') And ((ReklMemo.VorgangstypID) = 1 Or (ReklMemo.VorgangstypID) = 6 Or (ReklMemo.VorgangstypID) = 8)) " & _
                    "GROUP BY AufwandVorgangPositionen.Datum " & _
                    "HAVING (((AufwandVorgangPositionen.Datum)= #" & format(Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtDatum"), "mm\/dd\/yyyy") & "#));")



    If rsRekla.RecordCount Then
        Zeit = rsRekla("SummevonZeit")
        Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandReklamationen") = format(Zeit / 24) + 0
    Else
        Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandReklamationen") = 0
    End If

    Zeit = 0

    Set rsRep = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.Datum, Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit " & _
                "FROM (Mastertab LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID " & _
                "WHERE (((AufwandVorgangPositionen.Kürzel) = '" & get_pKuerzel() & "') And (ReklMemo.VorgangstypID) = 3 or (ReklMemo.VorgangstypID) = 7)" & _
                "GROUP BY AufwandVorgangPositionen.Datum " & _
                "HAVING (((AufwandVorgangPositionen.Datum)= #" & format(Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtDatum"), "mm\/dd\/yyyy") & "#));")

    If rsRep.RecordCount Then
        Zeit = rsRep("SummevonZeit")
        Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandRep") = format(Zeit / 24) + 0
    Else
        Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandRep") = 0
    End If

    Zeit = 0

    Set rsGU = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.Datum, Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit " & _
                "FROM (Mastertab LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID " & _
                "WHERE (((AufwandVorgangPositionen.Kürzel) = '" & get_pKuerzel() & "') And (ReklMemo.VorgangstypID) = 2)" & _
                "GROUP BY AufwandVorgangPositionen.Datum " & _
                "HAVING (((AufwandVorgangPositionen.Datum)= #" & format(Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtDatum"), "mm\/dd\/yyyy") & "#));")

    If rsGU.RecordCount Then
        Zeit = rsGU("SummevonZeit")
        Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandGU") = format(Zeit / 24) + 0
    Else
        Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandGU") = 0
    End If

    Forms("frmFormularROS").Form("frmZeiterfassungROS").Requery

    Set rsRekla = Nothing
    Set rsRep = Nothing
    Set rsGU = Nothing


    Dim Mittag As Double
    Dim Mittagspause As Date
    Dim SummeFehlzeit As Date

    If Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("chkMittagspause") = True Then
        Mittag = "0,75"
    Else
        Mittag = "0"
    End If

    Mittagspause = Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandReklamationen") + Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandGU") + Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtZeitaufwandRep") + Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtTätigkeitEinkauf") + format(Mittag / 24)

    SummeFehlzeit = Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtUhrzeitGehen") - Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtUhrzeitKommen") - Mittagspause

    If format((Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtUhrzeitGehen") - Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtUhrzeitKommen")), "hh:mm:ss") > Mittagspause Then
            Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtFehlzeit") = SummeFehlzeit
        Else
            Forms("frmFormularROS").Form("frmZeiterfassungROS").Controls("txtFehlzeit") = 0
    End If

    Me.Refresh

End Sub

Private Sub cmbUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbWreneingang_Click()

    DoCmd.OpenForm ("frmBestandServiceLager")

End Sub

Private Sub cmbWreneingang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me

End Sub

Private Sub cmbOffeneAB_Click()
On Error GoTo Err_cmbOffeneAB_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frmAB_Temp"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmbOffeneAB_Click:
    Exit Sub

Err_cmbOffeneAB_Click:
    MsgBox Err.Description
    Resume Exit_cmbOffeneAB_Click

End Sub

Private Sub Form_Timer()
'Aktuallisiert sich alle 5min, Zeitintervall steht auf 300000

    Me.frmGesperrteVorgänge.Requery

End Sub

Private Sub smbEmailDatenService_Click()

    DoCmd.OpenForm ("frmEmailServiceDaten")

End Sub
