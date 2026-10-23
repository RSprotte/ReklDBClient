Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
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

Private Sub FormSortieren(fsSpalte As String)
    Dim iResult As Integer
    iResult = MsgBox(getTC_FCFromCode("TC0302"), vbYesNoCancel)     'TC0302
    Select Case iResult
        Case vbYes
            Me.OrderBy = fsSpalte & ", ID"
        Case vbNo
            Me.OrderBy = fsSpalte & " DESC, ID DESC"
        Case vbCancel

    End Select
    Me.OrderByOn = True
    'Me.Requery
    Me.Refresh

End Sub

Private Sub Befehl60_Click()

If IsLoaded("Reklamationen_Übersicht_Konfigurator") Then
    Forms("Reklamationen_Übersicht_Konfigurator").SetFocus
Else
    DoCmd.OpenForm "Reklamationen_Übersicht_Konfigurator", acNormal, , , , acWindowNormal
    Forms("Reklamationen_Übersicht_Konfigurator").Controls("Bis") = Me.Bis
    Forms("Reklamationen_Übersicht_Konfigurator").Controls("Von") = Me.Von
    Forms("Reklamationen_Übersicht_Konfigurator").Controls("optKundenFilter") = 0
End If

    Me.cboFilterForm = ""

End Sub


Private Sub Bis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_NO)

End Sub

Private Sub cboFilterForm_Click() '20180905-ROS

    Dim ldb As DAO.Database
    Dim Date1 As String
    Dim Schaltjahr As String
    Dim strFormFilter As String
    Dim rsFormFilter As DAO.Recordset
    Dim ld_date As String
    Dim idKuerzel As Integer

    idKuerzel = DLookup("[SB-ID]", "Sachbearbeiter", "[Kurzzeichen] = '" & get_pKuerzel & "'")

    Set ldb = openAccDBConnection("openDB")
    Set rsFormFilter = ldb.OpenRecordset("Select * from tblFilterForm where [FilterBezeichnung] = '" & Me.cboFilterForm & "' And [FormName]= '" & Me.Name & "' And [SachbearbeiterID] = " & idKuerzel & " ;")

    strFormFilter = rsFormFilter("FilterParameter")

    If rsFormFilter("Dynamisch") Then
        strFormFilter = Replace(rsFormFilter("FilterParameter"), "#Von#", "#" & format(Me.Von, "mm\/dd\/yyyy") & "#")
        strFormFilter = Replace(strFormFilter, "#Bis#", "#" & format(Me.Bis, "mm\/dd\/yyyy") & "#")
    Else
        strFormFilter = rsFormFilter("FilterParameter")
        Me.Von = rsFormFilter("DatumVon")
        Me.Bis = rsFormFilter("DatumBis")
    End If
    DoCmd.OpenForm "Reklamationen_Übersicht", acNormal
    Forms("Reklamationen_Übersicht").Filter = strFormFilter ' & sAuftrag
    Forms("Reklamationen_Übersicht").FilterOn = True
    Forms("Reklamationen_Übersicht").FilterView = rsFormFilter("FilterParameter")
    Me.Requery
    Me.Refresh


'    DoCmd.Maximize

    Set rsFormFilter = Nothing

'    20180913-ROS
    Dim QS As DAO.QueryDef
    Set QS = ldb.QueryDefs("RSP_Ausw_Basisdaten_1_Export")
    If InStr(QS.sql, "Where") Then
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, "Where") + 5) & Forms(sÜbersicht).Filter = sFilter & ";"
        QS.sql = Left(QS.sql, InStr(QS.sql, "Where") + 5) & strFormFilter & ";"
    Else
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, ";") - 1) & " Where " & Forms(sÜbersicht).Filter = sFilter & ";"
        QS.sql = Left(QS.sql, InStr(QS.sql, ";") - 1) & " Where " & strFormFilter & ";"
    End If

End Sub




Private Sub cboZeiterfassung_Click()

    DoCmd.OpenForm ("frmZeiterfassung")

End Sub

Private Sub CmbhalbesJahr_Click()
    Const ExcelDateiName = "_Reklamationen_älter_halbes_Jahr.xlsx"  ' zu erzeugende Exceldatei
    Const tmpAbfrage = "qrytemp"                                   ' Abfrage wird automatisch erzeugt und gelöscht
    Dim ldb As DAO.Database
    Dim sSQLState  As String
    Dim qdf As DAO.QueryDef

    Datum = format(Now(), "yyyy/mm/dd_hh.mm")   ' 20160819-ROS Datum in Dateiname eingefügt

    sSQLState = "SELECT Mastertab.ID, Mastertab.Auftragsdatum, Mastertab.[Kategorie-Status], Mastertab.[Zuständiger Service-MA], Mastertab.Sachbearbeiter, [Artikel-all_avista_out].[LQ Art-Nr], tblFirma.Firma_Name " & _
                "FROM ((((Mastertab INNER JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID) LEFT JOIN Vorgangstyp ON ReklMemo.VorgangstypID = Vorgangstyp.VorgangstypID) INNER JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]) INNER JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]) INNER JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID " & _
                "WHERE (((Vorgangstyp.Anzeige)='Meldung') AND ((Mastertab.extern)=True))" & _
                "GROUP BY Mastertab.ID, Mastertab.Auftragsdatum, Mastertab.[Kategorie-Status], Mastertab.[Zuständiger Service-MA], Mastertab.Sachbearbeiter, [Artikel-all_avista_out].[LQ Art-Nr], tblFirma.Firma_Name " & _
                "HAVING (((Mastertab.Auftragsdatum) Between Date()-548 And Date()-183) AND ((Mastertab.[Kategorie-Status])<>'Abgeschlossen'))"


                Set ldb = openAccDBConnection("openDB")
                Set qdf = ldb.CreateQueryDef(tmpAbfrage, sSQLState)

                DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tmpAbfrage, Datum & ExcelDateiName, True
                DoCmd.DeleteObject acQuery, tmpAbfrage
                ShellExecute 0, "open", Datum & ExcelDateiName, "", "", 1

                Set rs = Nothing
End Sub

Private Sub BezACode_Click()

    FormSortieren "Auftragsdatum"

End Sub


Private Sub BezLieferant_Click()

    FormSortieren "Lieferantenname"

End Sub

Private Sub BezFehlerklasse_Click()

    FormSortieren "Kategorie-Fehleranalyse"

End Sub


Private Sub BezKunde_Click()

    FormSortieren "Firma"

End Sub

Private Sub BezMldgNr_Click()

    FormSortieren "[Meldungsnummer Kunde]"


End Sub

Private Sub BezQSVBearb_Click()

    FormSortieren "KZ-Sachbearbeiter"


End Sub

Private Sub BezQualifizierung_Click()

    FormSortieren "Kategorie-Problem"

End Sub

Private Sub BezReklNr_Click()

    FormSortieren "ID"

End Sub

Private Sub BezRueckDatum_Click()

    FormSortieren "Rückmeldung_Kunde_Verschiebung"

End Sub

Private Sub BezSaBearb_Click()

    FormSortieren "KZ-Service"


End Sub

Private Sub BezStatus_Click()

    FormSortieren "Kategorie-Status"

End Sub

Private Sub BezVerursacher_Click()

    FormSortieren "Verursacher"

End Sub


Private Sub cmbReklMappe_Click()

    Dim ReklnrKnd As String

    If IsNull(Me.Text42) Then
        ReklnrKnd = ""
    Else
        ReklnrKnd = Me.Text42
    End If

    Call VorgangsOrdner(Me.Text3, ReklnrKnd, Me.Text46, True)   '20190820-ROS


End Sub


Private Sub cmbRücklieferungKunde_Click()

    Call RücklieferungKunde

End Sub



Private Sub cmbRücklieferungLQ_Click()
    Const ExcelDateiName = "_RücklieferungLQ_älter_halbes_Jahr.xlsx"  ' zu erzeugende Exceldatei
    Const tmpAbfrage = "qrytemp"                                   ' Abfrage wird automatisch erzeugt und gelöscht
    Dim ldb As DAO.Database
    Dim sSQLState  As String
    Dim qdf As DAO.QueryDef

    Datum = format(Now(), "yyyy/mm/dd_hh.mm")   ' 20160819-ROS Datum in Dateiname eingefügt

    sSQLState = "SELECT Mastertab.ID, Mastertab.Auftragsdatum, tblFirma.Firma_Name, [Baugruppen-avista_out].[LQ Art-Nr] AS BG, [Artikel-all_avista_out]![LQ Art-Nr] & ' - ' & [Variante] AS Artikel, Mastertab.[Menge betroffen - Kunde] AS Stk " & _
                "FROM (((Mastertab INNER JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]) INNER JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID) " & _
                "LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]) " & _
                "LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID] " & _
                "WHERE (((Mastertab.[Kategorie-Status]) = 'Rücklieferung Kunde')) " & _
                "GROUP BY Mastertab.ID, Mastertab.Auftragsdatum, tblFirma.Firma_Name, [Baugruppen-avista_out].[LQ Art-Nr], [Artikel-all_avista_out]![LQ Art-Nr] & ' - ' & [Variante], Mastertab.[Menge betroffen - Kunde] " & _
                "HAVING (((Mastertab.Auftragsdatum) Between Date()-548 And Date()) AND ((tblFirma.Firma_Name)='LQCN' Or (tblFirma.Firma_Name)='LQUS'))"


                Set ldb = openAccDBConnection("openDB")
                Set qdf = ldb.CreateQueryDef(tmpAbfrage, sSQLState)

                DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tmpAbfrage, Datum & ExcelDateiName, True
                DoCmd.DeleteObject acQuery, tmpAbfrage
                ShellExecute 0, "open", Datum & ExcelDateiName, "", "", 1

                Set rs = Nothing

End Sub


Private Sub ExportExcel_Click()

    Dim strFilter As String

    ' Filter nur übergeben, wenn er im Formular auch gerade aktiv angewendet wird
    If Me.FilterOn Then
        strFilter = Me.Filter
    Else
        strFilter = ""
    End If

    ' Aufruf der globalen Prozedur mit Übergabe der Werte
    ' Optional können Sie als dritten Parameter einen eigenen Dateinamen mitgeben
    Call ExportRecordSourceNachExcel(Me.RecordSource, strFilter, Me.Name)

End Sub



Private Sub FilterLoeschen_Click()

    Dim Date1 As String
    Dim Schaltjahr As String

'    Date1 = "29.02."
'    Schaltjahr = Format(Date, "dd.mm.")
'    If Schaltjahr = Date1 Then
'        sVon = Format(Date - 1, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    Else
'        sVon = Format(Date, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    End If

    Date1 = "02/29/"
    Schaltjahr = format(Date, "mm\/dd\/")
    If Schaltjahr = Date1 Then
        sVon = format(Date - 1, "mm\/dd\/") & LTrim(Str(Val(format(Date, "yyyy")) - 1))
    Else
        sVon = format(Date, "mm\/dd\/") & LTrim(Str(Val(format(Date, "yyyy")) - 1))
    End If
    sBis = Date
    sAuftrag = " AND ([Auftragsdatum] between #" & format(sVon, "mm\/dd\/yyyy") & "# and #" & _
                                                format(sBis, "mm\/dd\/yyyy") & "#)"


'    Me.Von = sVon
    Me.Von = format(sVon, "mm\/dd\/yyyy")
    Me.Bis = format(Date, "dd\/mm\/yyyy")

'    Me.Filter = "([ErfasserIntern] Is Null)"
    Forms("Reklamationen_Übersicht").Filter = "([ErfasserIntern] Is Null" & sAuftrag & ")"
    Me.FilterView = Me.Filter
    Me.cboFilterForm = ""
    Me.Requery
'    Me.Refresh

End Sub

Private Sub FilterView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Activate()

'UpdateForm_Click
End Sub

Private Sub Form_Close()

On Error GoTo err_Form_Close

    DoCmd.Close acForm, "Reklamationen_Übersicht_Konfigurator"

err_Form_Close:
    Exit Sub

End Sub

Private Sub Form_DblClick(Cancel As Integer)
    Dim lsForm As String

    p_ReklNr = Me.ID
'    If Not IsNull(Me.ErfasserIntern) Then
'        OpenReklamationI
'        lsForm = "Fehlererfassung - Register"
'    Else
        NewReklamationE "Bearbeiten", get_p_ReklNr()
'        lsForm = "Fehlererfassung - Register"
'    End If
    Forms("Fehlererfassung - Register").SetFocus

end2_Form_DblClick:

End Sub

Private Sub Form_GotFocus()
'UpdateForm_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm


    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With



    Select Case Application.CurrentProject.Name
        Case "QS-FehlererfassungIntern.mdb"
            Me.Filter = "[intern] = True AND [ErfasserIntern] Is Not Null"
            Me.FilterOn = True
        Case Else
            Me.Filter = ""
    End Select

    setCmbViewFalse Me

    If get_ServiceZuordnung = True Then getCmbView Me, "Service"

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)


End Sub


Private Sub Form_Open(Cancel As Integer)

    Select Case Application.CurrentProject.Name
        Case "QS-FehlererfassungIntern.mdb"
            Me.Filter = "[intern] = True AND [ErfasserIntern] Is Not Null"
            Me.FilterOn = True
        Case Else
            Me.Filter = ""
    End Select
    Me.FilterView = Me.Filter

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Form_Timer()
    UpdateForm_Click
End Sub

Private Sub IntervallOFF_Click()

Me.TimerInterval = 0
Me.UpdateForm.SetFocus
Me.IntervallOFF.Enabled = False

End Sub



Private Sub Objekt_schließen___Formular_Click()
    FormKoorSave Me
    DoCmd.Close
End Sub


Private Sub Text86_DblClick(Cancel As Integer)

    p_ReklNr = Me.ID
    NewReklamationE "Bearbeiten", get_p_ReklNr()
    Forms("Fehlererfassung - Register").SetFocus

End Sub

Private Sub txtReklNrFilter_Change()
'On Error GoTo err_txtReklNrFilter_Change
'
'Dim ll_posIDFilter As Integer
'Dim ls_oldFilter As String
'
'ls_oldFilter = Me.Filter
'
'If IsNumeric(Right(Me.txtReklNrFilter.Properties("Text"), 1)) = False And Len(Me.txtReklNrFilter.Properties("Text")) > 0 And Len(Me.txtReklNrFilter.Properties("Text")) > Len(Me.txtReklNrFilter.OldValue) Then
'   Me.txtReklNrFilter.Text = Left(Me.txtReklNrFilter.Text, Len(Me.txtReklNrFilter.Text) - 1)
'Else
'
'    'Ist schon ein Filter gesetzt
'     If Len(Me.Filter) > 0 Then
'        'Filter gesetzt
'        'Ist Filter für Rekl-ID gesetzt
'        ll_posIDFilter = InStr(1, Me.Filter, "[ID]")
'        If ll_posIDFilter > 0 Then
'            'Filter Rekl-ID gesetzt
'            'Rekl-ID einzigster Filter
'            If ll_posIDFilter = 2 Then
'                'ReklID einzigster Filter
'                If Len(Me.txtReklNrFilter.Text) = 0 Then
'                    'Filter löschen
'                    Me.Filter = ""
'                    Me.FilterOn = False
'                Else
'                    'Filter anpassen
'                    Me.Filter = "([ID] like '" & Me.txtReklNrFilter.Text & "*')"
'                    Me.FilterOn = True
'                End If
'            Else
'                'ReklID letzter Eintrag im Filter
'                If Len(Me.txtReklNrFilter.Text) = 0 Or IsNull(Me.txtReklNrFilter) Then
'                    'Filter löschen
'                    ls_ReplString = MID(Me.Filter, ll_posIDFilter - 6, InStr(ll_posIDFilter, Me.Filter, "*") + 3 - (ll_posIDFilter - 6))
'                    Me.Filter = Replace(Me.Filter, ls_ReplString, "")
'                    Me.FilterOn = True
'                Else
'                    'Filter anpassen
'                    ls_ReplString = MID(Me.Filter, ll_posIDFilter + 11, InStr(ll_posIDFilter, Me.Filter, "*") - (ll_posIDFilter + 10))
'                    Me.Filter = Replace(Me.Filter, ls_ReplString, Me.txtReklNrFilter.Text & "*")
'                    Me.FilterOn = True
'                End If
'            End If
'        Else
'            'Filter vorhanden, aber Rekl-ID nicht -> Rekl-ID ergänzen
'            Me.Filter = Left(Me.Filter, Len(Me.Filter) - 1) & " AND ([ID] like '" & Me.txtReklNrFilter.Text & "*'))"
'            Me.FilterOn = True
'        End If
'    Else
'        If Len(Me.txtReklNrFilter.Text) > 0 Then
'            'Filter anpassen
'            Me.Filter = "([ID] like '" & Me.txtReklNrFilter.Text & "*')"
'            Me.FilterOn = True
'        End If
'    End If
'    Me.FilterView = Me.Filter
'    'Me.Requery
'    'Me.Refresh
'End If
'Me.txtReklNrFilter.SetFocus
'Me.txtReklNrFilter.SelStart = Len(Me.txtReklNrFilter.Properties("Text")) + 1
'Exit Sub
'
'err_txtReklNrFilter_Change:
'If Err = 2185 Then
'    Me.Filter = ls_oldFilter
'    Me.txtReklNrFilter.Properties("Text") = Left(Me.txtReklNrFilter.Text, Len(Me.txtReklNrFilter.Text) - 1)
'    Me.txtReklNrFilter.SelStart = Len(Me.txtReklNrFilter.Properties("Text")) + 1
'End If
    Dim strFilter As String         '20190823-ROS

    Dim intStart As Integer

    intStart = Me.txtReklNrFilter.SelStart

    If Not Len(Me.txtReklNrFilter.Text) = 0 Then
        strFilter = "[ID] like '" & Me.txtReklNrFilter.Text & "*'"
        Me.Filter = strFilter
        Me.FilterOn = True
        Me.txtReklNrFilter.SelStart = intStart
    Else
        Me.txtReklNrFilter = ""
        Me.FilterOn = False
        Me.txtReklNrFilter.SetFocus
    End If
End Sub

Private Sub UpdateForm_Click()
On Error GoTo err_UpdateForm_Click

Call OpenForm_Reklamationen_Übersicht

'If Screen.ActiveForm.Name = "Reklamationen_Übersicht" Then
'    Me.Refresh
'    Me.Requery
'End If
'
err_UpdateForm_Click:

End Sub

Private Sub Von_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_NO)

End Sub
Private Sub cboPersFilterSpeichern_Click() '20180905-ROS

'    Dim ldb As DAO.Database
'    Dim strFilterBezeichnung As String
'    Dim strFilterparameter As String
'
'    strFilterparameter = Me.Filter
'    'strFilterparameter = Me.FilterView.Value
'    strFilterparameter = Replace(strFilterparameter, "'", "''")
'
'    Set ldb = openAccDBConnection("openDB")
''    strFilterBezeichnung = InputBox("Bitte Filterbezeichnung eingeben")
''    strFilterBezeichnung = InputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
'    strFilterBezeichnung = getValueFromInputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
'
'    If strFilterBezeichnung = "" Then Exit Sub
'
'    li_retMsgBox = MsgBox(getTC_FCFromCode("TC4209"), vbYesNo + vbDefaultButton2)
'
'    If li_retMsgBox = 6 Then
'        ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
'                    "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & strFilterparameter & "', '" & strFilterBezeichnung & "' , True, " & _
'                    "False, #" & Format(Me.Von, "mm\/dd\/yyyy") & "#, #" & Format(Me.Bis, "mm\/dd\/yyyy") & "#);")
'    Else
'
'        ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
'                    "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & Replace(Me.FilterView, "'", "''") & "', '" & strFilterBezeichnung & "' , True, True, '','');")
'    End If

    Dim ldb As DAO.Database              '20260304-RSP
    Dim strFilterBezeichnung As String
    Dim strFilterparameter As String
    Dim llFilterId As Long

    strFilterparameter = Me.Filter
    strFilterparameter = Replace(strFilterparameter, "'", "''")

'    Set ldb = openAccDBConnection("openDB")
    strFilterBezeichnung = getValueFromInputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
    Set ldb = openAccDBConnection("openDB")

    If strFilterBezeichnung = "" Then Exit Sub

    li_retMsgBox = MsgBox(getTC_FCFromCode("TC4209"), vbYesNo + vbDefaultButton2)

    If li_retMsgBox = 6 Then
        llFilterId = Nz(DLookup("[FilterFormID]", "tblFilterForm", "[Filterbezeichnung]='" & strFilterBezeichnung & "' and [SachbearbeiterID] = " & get_SachbearbeiterID()), 0)
        If llFilterId = 0 Then
            ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
                        "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & strFilterparameter & "', '" & strFilterBezeichnung & "' , True, " & _
                        "False, #" & format(Me.Von, "mm\/dd\/yyyy") & "#, #" & format(Me.Bis, "mm\/dd\/yyyy") & "#);")
        Else
            'Filterbezeichnung schon vorhanden - Überschreiben?
            If MsgBox(getTC_FCFromCode("TC4281"), vbYesNo + vbDefaultButton1, getTC_FCFromCode("TC4282")) = vbYes Then
                ldb.Execute ("Update tblFilterForm set [FilterParameter] = '" & strFilterparameter & "', [DatumVon] = #" & format(Me.Von, "mm\/dd\/yyyy") & "#, " & _
                            "[DatumBis] = #" & format(Me.Bis, "mm\/dd\/yyyy") & "# " & _
                            "Where [FilterFormID] = " & llFilterId)
            End If
        End If
    Else
        llFilterId = Nz(DLookup("[FilterFormID]", "tblFilterForm", "[Filterbezeichnung]='" & strFilterBezeichnung & "' and [SachbearbeiterID] = " & get_SachbearbeiterID()), 0)
        If llFilterId = 0 Then
            'Set ldb = openAccDBConnection("openDB")
            ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
                        "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & Replace(Me.FilterView, "'", "''") & "', '" & strFilterBezeichnung & "' , True, True, '','');")
        Else
            'Filterbezeichnung schon vorhanden - Überschreiben?
            If MsgBox(getTC_FCFromCode("TC4281"), vbYesNo + vbDefaultButton1, getTC_FCFromCode("TC4282")) = vbYes Then
                ldb.Execute ("Update tblFilterForm set [FilterParameter] = '" & strFilterparameter & "', [DatumVon] = '', " & _
                            "[DatumBis] = '' " & _
                            "Where [FilterFormID] = " & llFilterId)
            End If
        End If
    End If


    Me.Requery
    Me.Refresh


End Sub
