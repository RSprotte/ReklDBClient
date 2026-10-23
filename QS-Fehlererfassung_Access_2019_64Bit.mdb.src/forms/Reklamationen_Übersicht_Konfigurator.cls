Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bTRUMPFGrp As Boolean  'Flag zur Einschränkung der Medlungsnummergröße von 7 Stellen, wenn über Button TRUMPF-GRUPPE alle TRUMPF Werke selektiert werden


Private Sub AlleKunden_Click()

Dim lngLoop As Long

Me.filKundenauswahl = ""
filKundenauswahl_Change

For lngLoop = 0 To Me.Kundenauswahl.ItemsSelected.Count - 1
Me.Kundenauswahl.Selected(Me.Kundenauswahl.ItemsSelected.Item(0)) = False
Next lngLoop

bTRUMPFGrp = False 'Einschränkung  Medlungslänge 7 Stellen zurückgesetzt

End Sub
Public Sub setValue(ByVal CntDate As Date)

    Me.ActiveControl = CntDate

End Sub
Private Sub Befehl48_Click()
    Me.OptionFilterauswahl = Null
End Sub



Private Sub Befehl61_Click()
Me.Baugruppe = Null
Me.Untergruppe = Null
Me.Artikel1 = Null
Me.KombAvistaArtikelTyp = Null

End Sub



Private Sub Bestellnummer_AfterUpdate()

    Me.OptionFilterauswahl = 7
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtFA_Nummer = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub Bis_Click()
    Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)



End Sub

Private Sub ButSelectKundeAuswListe_Click()

    Me.optKundenFilter = 0
    Me.Kundenauswahl.Locked = False

End Sub

Private Sub ButtonAlleErfasser_Click()

Me.KombEinheitErfasser = Null
Me.KombEinheitErfasser.Requery
Me.KombBereichErfasser = Null
Me.KombBereichErfasser.Requery
Me.KombFunktionErfasser = Null
Me.KombFunktionErfasser.Requery
Me.KombWorkstationErfasser = Null
Me.KombWorkstationErfasser.Requery
Me.KombWorkstationErfasser.Enabled = True

End Sub

Private Sub ErfasserIntern_AfterUpdate()

Me.OptionFilterauswahl = 6
Me.Meldungsnummer = Null
Me.txtRMA = Null
Me.Bestellnummer = Null
Me.KombiServiceMA = Null
Me.KombiSachbearbeiter = Null
Me.KombiAnsprechpartner = Null
Me.KombiLieferant = Null
Me.txtFA_Nummer = Null
Me.txtRepBestellnummer = Null
Me.txtMaschNr = Null

End Sub

Private Sub Fehlerquelle_AfterUpdate()

Select Case Me.Fehlerquelle
    Case 1:
        Me.MeldenummerVorhanden = False
    Case 2:
    Case 3:
End Select

End Sub

Private Sub filKundenauswahl_Change()
On Error GoTo err_filKundenauswahl_Change

    Dim lsFilter As String

    If Me.filKundenauswahl.Text = "" Then
        Me.Kundenauswahl.RowSource = "SELECT tblFirma.Firma_Name From tblFirma GROUP BY tblFirma.Firma_Name;"
    Else
        Me.Kundenauswahl.RowSource = "SELECT tblFirma.Firma_Name From tblFirma Where [Firma_Name] like '" & Me.filKundenauswahl.Text & "*' GROUP BY tblFirma.Firma_Name;"
    End If
    Me.Kundenauswahl.Requery
    Me.filKundenauswahl.SetFocus
    Me.filKundenauswahl.SelStart = Len(Me.filKundenauswahl.Text)

    Exit Sub

err_filKundenauswahl_Change:

    If Err = 2185 Then
        Me.filKundenauswahl.SetFocus
        Resume
    End If




End Sub

Private Sub FilterVerursacher_Click()

If Me.FilterVerursacher = "Lieferant" Then
    Me.FilterVerantwSachbLief.Visible = True
    Me.FilterLiefVerantwLQ.Visible = True
    Me.FilterLiefVerantwLQ.Value = Null
Else
    Me.FilterVerantwSachbLief = False
    Me.FilterLiefVerantwLQ.Visible = False
End If
    Me.FilterVerantwSachbLief = Null  '20160630ROS
End Sub

Private Sub Form_Activate()
    'DoCmd.Maximize
End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm


    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With
'    Me.Von.Value = DateAdd("yyyy", -1, Date)      '20160627-ROS

End Sub

Private Sub Form_Open(Cancel As Integer)

Dim sRowSource As String

Me.OptionFilterauswahl = Null
Me.PEbezogen = False

sRowSource = "SELECT [Lieferanten].[Lieferant-ID], [Lieferanten].[Lieferantenname] " & _
             "FROM Lieferanten ORDER BY [Lieferanten].[Lieferantenname];"

Me.KombiLieferant.RowSource = sRowSource
Me.KombiLieferant.Requery
'Reklamationsauswahl festlegen

Select Case Application.CurrentProject.Name
    Case "QS-FehlererfassungEntwicklung.mdb"
    Case "QS-FehlererfassungExtern.mdb"
    Case "QS-FehlererfassungIntern.mdb"
        Me.Fehlerquelle = 1
        Me.Fehlerquelle.Locked = True
End Select
'Init Kundenauswahl TRUMPF GRUPPE
bTRUMPFGrp = False

'20160630-ROS
    Me.FilterLiefVerantwLQ.Visible = False
    Me.FilterVerantwSachbLief.Visible = False
    Me.optKundenFilter = 0
    Me.Refresh

    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub


Private Sub KeinFilterVerursacher_Click()

Me.FilterVerursacher = Null
If Me.FilterVerursacher = "Lieferant" Then
    Me.FilterVerantwSachbLief.Visible = True
    Me.FilterLiefVerantwLQ.Visible = True
    Me.FilterLiefVerantwLQ.Value = Null
Else
    Me.FilterVerantwSachbLief.Visible = False
    Me.FilterLiefVerantwLQ.Visible = False
End If

End Sub

Private Sub KlassifizierungAlle_Click()
Me.ListTaetigkeiten = Null
Me.ListFehlerbild.Requery
End Sub

Private Sub KombBereichErfasser_AfterUpdate()

Me.KombWorkstationErfasser = Null
Me.KombWorkstationErfasser.Requery
Me.KombFunktionErfasser = Null
Me.KombFunktionErfasser.Requery

If IsNull(Me.KombBereichErfasser) Then
    Me.KombFunktionErfasser = Null
    Me.KombFunktionErfasser.Requery
    Me.KombFunktionErfasser.Enabled = True
Else
    Me.KombFunktionErfasser = Null
    Me.KombFunktionErfasser.Requery
    Me.KombFunktionErfasser.Enabled = False
End If

End Sub

Private Sub KombEinheitErfasser_AfterUpdate()

'Nach aktualisierung werden alle abhängigen Felder initialisiert und angezeigt
Me.KombBereichErfasser = Null
Me.KombBereichErfasser.Requery
Me.KombBereichErfasser.Enabled = True
Me.KombWorkstationErfasser = Null
Me.KombWorkstationErfasser.Requery
Me.KombWorkstationErfasser.Enabled = True
Me.KombFunktionErfasser = Null
Me.KombFunktionErfasser.Requery
Me.KombFunktionErfasser.Enabled = True


End Sub

Private Sub KombFunktionErfasser_AfterUpdate()

If IsNull(Me.KombFunktionErfasser) Then
    Me.KombWorkstationErfasser = Null
    Me.KombWorkstationErfasser.Requery
    Me.KombWorkstationErfasser.Enabled = True
    Me.KombBereichErfasser = Null
    Me.KombBereichErfasser.Requery
    Me.KombBereichErfasser.Enabled = True

Else
    Me.KombWorkstationErfasser = Null
    Me.KombWorkstationErfasser.Requery
    Me.KombWorkstationErfasser.Enabled = False
    Me.KombBereichErfasser = Null
    Me.KombBereichErfasser.Requery
    Me.KombBereichErfasser.Enabled = False
End If
End Sub

Private Sub KombiAnsprechpartner_AfterUpdate()
Me.OptionFilterauswahl = 4
Me.Meldungsnummer = Null
Me.txtRMA = Null
Me.Bestellnummer = Null
Me.KombiServiceMA = Null
Me.KombiSachbearbeiter = Null
Me.ErfasserIntern = Null
Me.KombiLieferant = Null
Me.txtFA_Nummer = Null
Me.txtRepBestellnummer = Null
Me.txtMaschNr = Null

End Sub

Private Sub KombiLieferant_AfterUpdate()

    Me.OptionFilterauswahl = 5
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.txtFA_Nummer = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub KombiSachbearbeiter_AfterUpdate()

    Me.OptionFilterauswahl = 3
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtFA_Nummer = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub KombiServiceMA_AfterUpdate()

    Me.OptionFilterauswahl = 2
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtFA_Nummer = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub ListTaetigkeiten_AfterUpdate()
Me.ListFehlerbild.Requery
End Sub

Private Sub MatchLieferant_AfterUpdate()
Dim sRowSource As String
sRowSource = "SELECT [Lieferanten].[Lieferant-ID], [Lieferanten].[Lieferantenname] " & _
             "FROM Lieferanten ORDER BY [Lieferanten].[Lieferantenname];"

Me.KombiLieferant.RowSource = Left(sRowSource, InStr(sRowSource, "ORDER BY") - 1) & _
                              "Where [Lieferanten].[Lieferantenname] like '" & Me.MatchLieferant & "*' " & _
                              Right(sRowSource, Len(sRowSource) - InStr(sRowSource, "ORDER BY") + 1)
Me.KombiLieferant.Requery
End Sub

Private Sub Meldungsnummer_AfterUpdate()

    Me.OptionFilterauswahl = 1
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtFA_Nummer = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub OpenDatumVonBis_Click()
DoCmd.OpenForm "DatumVonBis"
End Sub


Private Sub optKundenFilter_AfterUpdate()

    If Me.optKundenFilter > 0 Then
        AlleKunden_Click
        Me.Kundenauswahl.Locked = True
    End If

End Sub

Private Sub QualifizierungAlle_Click()
Me.Qualifizierung = Null
End Sub

Private Sub StatusAlle_Click()
Me.Status = Null
End Sub



Private Sub TagAlle_Click()

Me.ListTag = Null

End Sub


Private Sub txtFA_Nummer_AfterUpdate()

    Me.OptionFilterauswahl = 9
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub txtMaschNr_AfterUpdate()

    Me.OptionFilterauswahl = 11
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtRepBestellnummer = Null
    Me.txtFA_Nummer = Null

End Sub

Private Sub txtRepBestellnummer_AfterUpdate()

    Me.OptionFilterauswahl = 10
    Me.Meldungsnummer = Null
    Me.txtRMA = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtFA_Nummer = Null
    Me.txtMaschNr = Null


End Sub

Private Sub txtRMA_AfterUpdate()

    Me.OptionFilterauswahl = 8
    Me.Meldungsnummer = Null
    Me.Bestellnummer = Null
    Me.KombiServiceMA = Null
    Me.KombiSachbearbeiter = Null
    Me.ErfasserIntern = Null
    Me.KombiAnsprechpartner = Null
    Me.KombiLieferant = Null
    Me.txtFA_Nummer = Null
    Me.txtRepBestellnummer = Null
    Me.txtMaschNr = Null

End Sub

Private Sub Übernehmen_Click()
On Error GoTo err_Übernehmen_Click

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sÜbersicht As String
    Dim ls_DateVon As String
    Dim ls_DateBis As String
    Dim lb_ADate As Boolean

'    sÜbersicht = "Reklamationen_Übersicht"
    bFilterOn = False
    sfilter = ""
    sKndAuswQuery = csKundenAuswahl
    lb_ADate = False

    '20210407-ROS Auswahl des aktuellen Übersichts-Fenster
    If IsLoaded("Reklamationen_ÜbersichtIntern") Then
        sÜbersicht = "Reklamationen_ÜbersichtIntern"
    End If

    If IsLoaded("Reklamationen_Übersicht") Then
        sÜbersicht = "Reklamationen_Übersicht"
    End If



    sfilter = "("
    'RSP - 2012.06.21 Einführung Filter FilterVerursacher
    'RSP - 2012.06.25 Fehlerbearbeitung
    'RSP - 2013.10.01 Einführung Filter FilterVerantwSachbLief und FilterLiefVerantwLQ

    If Not IsNull(Me.FilterVerursacher) Then
        sfilter = sfilter + "([Verursacher] = '" & Me.FilterVerursacher & "')"
        bFilterOn = True

        If Me.FilterVerursacher = "Lieferant" Then
            If Not IsNull(Me.FilterVerantwSachbLief) Then
                sfilter = sfilter + "And ([VerantwortlichEinkauf] = '" & Me.FilterVerantwSachbLief & "')"
            End If
            If Not IsNull(Me.FilterLiefVerantwLQ) Then
                sfilter = sfilter + "And ([LieferantenverantwortungLQ] = " & Me.FilterLiefVerantwLQ & ")"

            End If
        End If
    End If
    'RSP - 2012.06.25 Fehlerbearbeitung
    If Not IsNull(Me.OptionFilterauswahl) And (Me.OptionFilterauswahl > 0) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If
        Select Case Me.OptionFilterauswahl
            Case 1  'Meldungsnummer
                sfilter = sfilter + "([Meldungsnummer Kunde] like '" & Me.Meldungsnummer & "*')"
            Case 8  'RMA
                sfilter = sfilter + "([Zuständig] like '" & Me.txtRMA & "*')"
            Case 7  'Bestellnummer
                sfilter = sfilter + "([Bestellnummer Kunde] like '" & Me.Bestellnummer & "*')"
            Case 2  'Service-MA
                sfilter = sfilter + "([Zuständiger Service-MA] = '" & Me.KombiServiceMA & "')"
            Case 3  'Sachbearbeiter
                If Me.PEbezogen Then
                    Set ldb = openAccDBConnection("openDB")
                    Set rs = ldb.OpenRecordset("SELECT Sachbearbeiter.[SB-ID], Sachbearbeiter.Einheit, Sachbearbeiter.Name, LQEinheit.ID " & _
                                              "FROM Sachbearbeiter INNER JOIN LQEinheit ON Sachbearbeiter.Einheit = LQEinheit.LQEinheit " & _
                                              "where Name = '" & _
                                               Me.KombiSachbearbeiter & "';")
                    sfilter = sfilter + "([LQEinheitID] = " & rs("ID").Value & ")"
                    Set rs = Nothing
                    'closeAccDBConnection db
                Else
                    sfilter = sfilter + "([Sachbearbeiter] = '" & Me.KombiSachbearbeiter & "')"
                End If
            Case 6  'ErfasserIntern
                sfilter = sfilter + "([ErfasserIntern] = " & Me.ErfasserIntern & ")"
            Case 4  'Ansprechpartner Kunde
                sfilter = sfilter + "([Kontaktperson] = '" & Me.KombiAnsprechpartner & "')"
            Case 5  'Lieferant
                sfilter = sfilter + "([Lieferanten-Code 1] = " & Me.KombiLieferant & ")"
            Case 9
                sfilter = sfilter + "([Fertigungsauftrag-FA] = '" & Me.txtFA_Nummer & "')"
            Case 10 'Bestellnummer Rep.Auftrag      20190613-ROS
                sfilter = sfilter + "([BestellnrRepAuftrag] like '" & Me.txtRepBestellnummer & "*')"
            Case 11 'Nach MaschNr. suchen            20250722-ROS
                sfilter = sfilter + "([Maschinennummer] like '" & Me.txtMaschNr & "*')"
        End Select
        bFilterOn = True
    End If
    If Me.AlleReklamationen2 Then
        If bFilterOn Then
            sfilter = sfilter + "And ([Kategorie-Status] <> 'Abgeschlossen')"
        Else
            sfilter = sfilter + "([Kategorie-Status] <> 'Abgeschlossen')"
        End If
        bFilterOn = True
    End If

    bKunde = False
'20260723-ROS: Änderungen wegen neuer Kundengruppen in Avista
'#############################################################
    If Me.optKundenFilter Then
        Select Case Me.optKundenFilter
                Case 1 'Trumpf-Gruppe
                    sKunde = 4
                Case 2 'sonstige Kunden
                    sKunde = 0
                Case 3 'LQ
                    sKunde = 7
        End Select

'        If bFilterOn Then
'            sfilter = sfilter + "And ([Zuordnung_ID] = " & sKunde & ")"
'        Else
'            sfilter = sfilter + "([Zuordnung_ID] = " & sKunde & ")"
'        End If

        If bFilterOn Then
            If sKunde = 0 Then
                sKunde = 1
                sfilter = sfilter + "And ([Zuordnung_ID] <> " & sKunde & ")"
            Else
                sfilter = sfilter + "And ([Zuordnung_ID] = " & sKunde & ")"
            End If
        Else
            If sKunde = 0 Then
                sKunde = 4
                sfilter = sfilter + "([Zuordnung_ID] <> " & sKunde & ")"
            Else
                sfilter = sfilter + "([Zuordnung_ID] = " & sKunde & ")"
            End If
        End If

        bFilterOn = True
        bKunde = True
    End If

    If Me.Kundenauswahl.ItemsSelected.Count And bKunde = False Then
        If bFilterOn Then
            sfilter = sfilter & " And ([Firma_Name] in ("
        Else
            sfilter = sfilter & "([Firma_Name] in ("
            bFilterOn = True
        End If
        For Each varElem In Me.Kundenauswahl.ItemsSelected
            sfilter = sfilter & "'" & Me.Kundenauswahl.ItemData(varElem) & "',"
        Next varElem
        sfilter = Left(sfilter, Len(sfilter) - 1) & "))"
        bFilterOn = True
    End If

    If Me.Status.ItemsSelected.Count Then
        If bFilterOn Then
            sfilter = sfilter & " And ([Kategorie-Status] in ("
        Else
            sfilter = sfilter & "([Kategorie-Status] in ("
        End If
        For Each varElem In Me.Status.ItemsSelected
            sfilter = sfilter & "'" & Me.Status.ItemData(varElem) & "',"
        Next varElem
        sfilter = Left(sfilter, Len(sfilter) - 1) & "))"
        bFilterOn = True
    End If

'20160128-ROS
    If Me.ListTaetigkeiten.ItemsSelected.Count Then
        If bFilterOn Then
            Select Case Me.OptionFehlerort
                Case 1 'Erfasser
                    sfilter = sfilter & " And [TaetigkeitErfasserID] in (XXXXX)"
                Case 2 'Analyse
                    sfilter = sfilter & " And [TaetigkeitID] in (XXXXX)"
                Case 3 'Beides
                    sfilter = sfilter & " And ([TaetigkeitID]  in (XXXXX) OR [TaetigkeitErfasserID] in (XXXXX))"
            End Select
'
'            sFilter = sFilter & " And ([TaetigkeitID] in ("
        Else
            Select Case Me.OptionFehlerort
                Case 1 'Erfasser
                    sfilter = sfilter & "([TaetigkeitErfasserID] in (XXXXX))"
                Case 2 'Analyse
                    sfilter = sfilter & "([TaetigkeitID] in (XXXXX))"
                Case 3 'Beides
                    sfilter = sfilter & "([TaetigkeitID]  in (XXXXX) OR [TaetigkeitErfasserID] in (XXXXX))"
            End Select
'            sFilter = sFilter & "([TaetigkeitID] in ("
        End If
        sItemList = ""
        For Each varElem In Me.ListTaetigkeiten.ItemsSelected
            sItemList = sItemList & Me.ListTaetigkeiten.ItemData(varElem) & ","
        Next varElem
        sItemList = Left(sItemList, Len(sItemList) - 1)
        sfilter = Replace(sfilter, "XXXXX", sItemList)
        bFilterOn = True
    End If

    'Filter Fehlerbild kann nur gesetzt werden wenn nur eine Tätigkeit ausgewählt wurde
    If Me.ListFehlerbild.ItemsSelected.Count And Me.ListTaetigkeiten.ItemsSelected.Count = 1 Then
        If bFilterOn Then
            Select Case Me.OptionFehlerort
                Case 1 'Erfasser
                    sfilter = sfilter & " And [DefTaetigkeitErfasserID] in (XXXXX)"
                Case 2 'Analyse
                    sfilter = sfilter & " And [DefTaetigkeitID] in (XXXXX)"
                Case 3 'Beides
                    sfilter = sfilter & " And ([DefTaetigkeitID]  in (XXXXX) OR [DefTaetigkeitErfasserID] in (XXXXX))"
            End Select
'            sFilter = sFilter & " And ([DefTaetigkeitID]in ("
        Else
            Select Case Me.OptionFehlerort
                Case 1 'Erfasser
                    sfilter = sfilter & "([DefTaetigkeitErfasserID] in (XXXXX))"
                Case 2 'Analyse
                    sfilter = sfilter & "([DefTaetigkeitID] in (XXXXX))"
                Case 3 'Beides
                    sfilter = sfilter & "([DefTaetigkeitID]  in (XXXXX) OR [DefTaetigkeitErfasserID] in (XXXXX))"
            End Select
'            sFilter = sFilter & "([DefTaetigkeitID]in ("
            bFilterOn = True
        End If
        sItemList = ""
        For Each varElem In Me.ListFehlerbild.ItemsSelected
            sItemList = sItemList & Me.ListFehlerbild.ItemData(varElem) & ","
        Next varElem
        sItemList = Left(sItemList, Len(sItemList) - 1)
        sfilter = Replace(sfilter, "XXXXX", sItemList)
        bFilterOn = True
    End If

    If Me.ListTag.ItemsSelected.Count Then
        If bFilterOn Then
            sfilter = sfilter & " And ([TagID] in ("
        Else
            sfilter = sfilter & "([TagID] in ("
        End If
        For Each varElem In Me.ListTag.ItemsSelected
            sfilter = sfilter & Me.ListTag.ItemData(varElem) & ","
        Next varElem
        sfilter = Left(sfilter, Len(sfilter) - 1) & "))"
        bFilterOn = True
    End If

    '### Filter setzen zur Fehlerortbestimmung
    If IsNull(Me.KombWorkstationErfasser) Then
        If IsNull(Me.KombFunktionErfasser) Then
            If IsNull(Me.KombBereichErfasser) Then
                If Not IsNull(Me.KombEinheitErfasser) Then
                    If bFilterOn Then
                        sfilter = sfilter & " And ("
                    Else
                        sfilter = sfilter & "("
                    End If
                    Select Case Me.OptionFehlerort
                        Case 1 'Erfasser
                            sfilter = sfilter & "[LQEinheitErfasserID] = " & Me.KombEinheitErfasser & ")"
                        Case 2 'Analyse
                            sfilter = sfilter & "[LQEinheitID] = " & Me.KombEinheitErfasser & ")"
                        Case 3 'Beides
                            sfilter = sfilter & "([LQEinheitID] = " & Me.KombEinheitErfasser & ") OR ([LQEinheitErfasserID] = " & Me.KombEinheitErfasser & "))"
                    End Select
                    bFilterOn = True
                End If
            Else
                If bFilterOn Then
                    sfilter = sfilter & " And ("
                Else
                    sfilter = sfilter & "("
                End If
                Select Case Me.OptionFehlerort
                    Case 1 'Erfasser
                        sfilter = sfilter & "[LQBereichErfasserID] = " & Me.KombBereichErfasser & ")"
                    Case 2 'Analyse
                        sfilter = sfilter & "[LQBereichID] = " & Me.KombBereichErfasser & ")"
                    Case 3 'Beides
                        sfilter = sfilter & "([LQBereichID] = " & Me.KombBereichErfasser & ") OR ([LQBereichErfasserID] = " & Me.KombBereichErfasser & "))"
                End Select
                bFilterOn = True
            End If
        Else
            If bFilterOn Then
                sfilter = sfilter & " And ("
            Else
                sfilter = sfilter & "("
            End If
            Select Case Me.OptionFehlerort  '20180213-ROS Filteranpasung für Filterung der Funktionen
                Case 1 'Erfasser
                    sfilter = sfilter & "[LQWorkstationErfasserID] IN (SELECT [ID-WS] FROM LQWSzuGrp WHERE [ID-LQFunktion] = " & Me.KombFunktionErfasser & "))"
                Case 2 'Analyse
                    sfilter = sfilter & "[LQWorkstationID] IN (SELECT [ID-WS] FROM LQWSzuGrp WHERE [ID-LQFunktion] = " & Me.KombFunktionErfasser & "))"
                Case 3 'Beides
                    sfilter = sfilter & "([LQWorkstationErfasserID] IN (SELECT [ID-WS] FROM LQWSzuGrp WHERE [ID-LQFunktion] = " & Me.KombFunktionErfasser & ") OR " & _
                                        " [LQWorkstationID] IN (SELECT [ID-WS] FROM LQWSzuGrp WHERE [ID-LQFunktion] = " & Me.KombFunktionErfasser & ")))"
            End Select
            bFilterOn = True
        End If
    Else
        If bFilterOn Then
            sfilter = sfilter & " And ("
        Else
            sfilter = sfilter & "("
        End If
        Select Case Me.OptionFehlerort
            Case 1 'Erfasser
                sfilter = sfilter & "[LQWorkstationErfasserID] = " & Me.KombWorkstationErfasser & ")"
            Case 2 'Analyse
                sfilter = sfilter & "[LQWorkstationID] = " & Me.KombWorkstationErfasser & ")"
            Case 3 'Beides
                sfilter = sfilter & "([LQWorkstationID] = " & Me.KombWorkstationErfasser & ") OR ([LQWorkstationErfasserID] = " & Me.KombWorkstationErfasser & "))"
        End Select
        bFilterOn = True
    End If


    bIsNull = False 'Merker Item "--NULL--" nicht gesetzt
    If Me.Qualifizierung.ItemsSelected.Count Then
        Dim sFilterTemp As String
        sFilterTemp = ""
        For Each varElem In Me.Qualifizierung.ItemsSelected
            If Me.Qualifizierung.ItemData(varElem) = "--Null--" Then
                bIsNull = True
            Else
                sFilterTemp = sFilterTemp & "'" & Me.Qualifizierung.ItemData(varElem) & "',"
            End If
        Next varElem
        If bFilterOn Then
            sfilter = sfilter & " And "
        End If

        If Len(sFilterTemp) > 0 Then
            If bIsNull Then
                sfilter = sfilter & "([Kategorie-Problem] in (" & Left(sFilterTemp, Len(sFilterTemp) - 1) & ") or [Kategorie-Problem] is Null)"
            Else
                sfilter = sfilter & "([Kategorie-Problem] in (" & Left(sFilterTemp, Len(sFilterTemp) - 1) & "))"
            End If
        Else
            If bIsNull Then
                sfilter = sfilter & "([Kategorie-Problem] is Null)"
            End If
        End If
        bFilterOn = True
    End If


    If Me.RueckmeldedatumAnKunde Then
        If Not bFilterOn Then
            sfilter = sfilter + "([Rückmeldung_Kunde_Soll] Is not Null)"
            bFilterOn = True
        Else
            sfilter = sfilter + " And ([Rückmeldung_Kunde_Soll] Is not Null)"
        End If
    End If

    If Me.MeldenummerVorhanden = True Then
        If Not bFilterOn Then
            If bTRUMPFGrp Then
                sfilter = sfilter + "([Meldungsnummer Kunde] like '#######*')"
            Else
                sfilter = sfilter + "([Meldungsnummer Kunde] Is not Null)"
            End If
            bFilterOn = True
        Else
            If bTRUMPFGrp Then
                sfilter = sfilter + " And ([Meldungsnummer Kunde] like '#######*')"
            Else
                sfilter = sfilter + " And ([Meldungsnummer Kunde] Is not Null)"
            End If
        End If
    End If

    Select Case Me.Fehlerquelle

        Case 2:
            If Not bFilterOn Then
                sfilter = sfilter + "([extern] = True)"
                bFilterOn = True
            Else
                sfilter = sfilter + " And ([extern] = True)"
            End If
        Case 1:
            If Not bFilterOn Then
                sfilter = sfilter + "([intern] = True AND [ErfasserIntern] Is not Null)"
                bFilterOn = True
            Else
                sfilter = sfilter + " And ([intern] = True AND [ErfasserIntern] Is not Null)"
            End If
        Case 4:
        Dim Lieferant As String
            Lieferant = getTextFromLokalisierungsTab("TableItem", "Verursacher", 1779)
            If Not bFilterOn Then
'                sFilter = sFilter + "([intern] = True AND [ErfasserIntern] Is Null)"
                sfilter = sfilter + "([intern] = True AND[Verursacher] = '" & Lieferant & "')"
                bFilterOn = True
            Else
                sfilter = sfilter + " And ([intern] = True AND[Verursacher] = '" & Lieferant & "')"
            End If
    End Select

    If Me.ListVorgangstyp.ItemsSelected.Count Then
        If bFilterOn Then
            sfilter = sfilter & " And ([VorgangstypID] in ("
        Else
            sfilter = sfilter & "([VorgangstypID] in ("
        End If
        For Each varElem In Me.ListVorgangstyp.ItemsSelected
            sfilter = sfilter & Me.ListVorgangstyp.ItemData(varElem) & ","
        Next varElem
        sfilter = Left(sfilter, Len(sfilter) - 1) & "))"
        bFilterOn = True
    End If

    If Not ((Me.Baugruppe = "") Or (IsNull(Me.Baugruppe))) Then
        If bFilterOn Then
            sfilter = sfilter & " And ([Baugruppen-avista_out].[LQ Art-Nr] like '" & Me.Baugruppe & "*')"
        Else
            sfilter = sfilter & "([Baugruppen-avista_out].[LQ Art-Nr] like '" & Me.Baugruppe & "*')"
        End If
        bFilterOn = True
    End If

    If Not ((Me.Untergruppe = "") Or (IsNull(Me.Untergruppe))) Then
        If bFilterOn Then
            sfilter = sfilter & " And ([LINK-Nr - UBG] like '" & Left(Me.Untergruppe, 7) & "*')"
        Else
            sfilter = sfilter & "([LINK-Nr - UBG] like '" & Left(Me.Untergruppe, 7) & "*')"
        End If
        bFilterOn = True
    End If

    If Not ((Me.Artikel1 = "") Or (IsNull(Me.Artikel1))) Then
        If bFilterOn Then
            sfilter = sfilter & " And ( [Artikel-all_avista_out].[LQ Art-Nr] like '" & Me.Artikel1 & "*')"
        Else
            sfilter = sfilter & "( [Artikel-all_avista_out].[LQ Art-Nr] like '" & Me.Artikel1 & "*')"
        End If
        bFilterOn = True
    End If

    If Not ((Me.KombAvistaArtikelTyp = "") Or (IsNull(Me.KombAvistaArtikelTyp))) Then
        If bFilterOn Then
            sfilter = sfilter & " And ([AvistaArtikelTyp] = " & Me.KombAvistaArtikelTyp & ")"
        Else
            sfilter = sfilter & "([AvistaArtikelTyp] = " & Me.KombAvistaArtikelTyp & ")"
        End If
        bFilterOn = True
    End If

    If Not ((Me.Von = "")) Then
        ls_DateVon = format(Me.Von, "mm\/dd\/yyyy")
        ls_DateBis = format(Me.Bis, "mm\/dd\/yyyy")
        'lb_ADate = True
        If Me.Bis = "" Or IsNull(Me.Bis) Then
            Me.Bis = Date
            Me.Bis.Requery
        End If
        If bFilterOn Then
'        If lb_ADate And bFilterOn = True Then       '20181008-ROS: And bFilterOn eingefügt
'            sFilter = sFilter & " And ([Auftragsdatum] between #" & Format(Me.Von, "mm\/dd\/yyyy") & "# and #" & _
'                                                Format(Me.Bis, "mm\/dd\/yyyy") & "#)"
            sfilter = sfilter & " And ([Auftragsdatum] between #Von# and #Bis#)"
        Else
'            sFilter = sFilter & "([Auftragsdatum] between #" & Format(Me.Von, "mm\/dd\/yyyy") & "# and #" & _
'                                                Format(Me.Bis, "mm\/dd\/yyyy") & "#)"
            sfilter = sfilter & "([Auftragsdatum] between #Von# and #Bis#)"
            bFilterOn = True
        End If
        Select Case sÜbersicht
            Case "Reklamationen_Übersicht"
                Forms("Reklamationen_Übersicht").Controls("Von") = Me.Von
                Forms("Reklamationen_Übersicht").Controls("Bis") = Me.Bis

            Case "Reklamationen_ÜbersichtIntern"
                Forms("Reklamationen_ÜbersichtIntern").Controls("Von") = Me.Von
                Forms("Reklamationen_ÜbersichtIntern").Controls("Bis") = Me.Bis
        End Select
    End If



    If Not bFilterOn Then
        sfilter = ""
    Else
        sfilter = sfilter + ")"
    End If

    Select Case sÜbersicht
        Case "Reklamationen_Übersicht"
            If Not IsLoaded(sÜbersicht) Then
                DoCmd.OpenForm sÜbersicht
            End If
            If IsLoaded("Reklamationen_ÜbersichtIntern") Then
                DoCmd.Close acForm, "Reklamationen_ÜbersichtIntern"
            End If

        Case "Reklamationen_ÜbersichtIntern"
            If Not IsLoaded(sÜbersicht) Then
                DoCmd.OpenForm sÜbersicht
            End If
            If IsLoaded("Reklamationen_Übersicht") Then
                DoCmd.Close acForm, "Reklamationen_Übersicht"
            End If
    End Select
    If Not IsLoaded(sÜbersicht) Then
        DoCmd.OpenForm sÜbersicht
    End If
    'Filter in Übersicht setzen und aktualisieren
    Forms(sÜbersicht).Controls("FilterView") = sfilter

'    MsgBox "Länge Filter: " & Len(sFilter)
'    sFilter = Replace(sFilter, "#Von#", "#" & ls_DateVon & "#")
'    sFilter = Replace(sFilter, "#Bis#", "#" & ls_DateBis & "#")
    sfilter = Replace(sfilter, "Von", ls_DateVon)
    sfilter = Replace(sfilter, "Bis", ls_DateBis)
    Forms(sÜbersicht).Filter = sfilter
    Forms(sÜbersicht).FilterOn = True
    Forms(sÜbersicht).Requery
    Forms(sÜbersicht).Refresh

    'Query zum Exportieren Filter setzen (Where)
    'Dim ldb As DAO.Database
    Dim QS As DAO.QueryDef
    Set ldb = openAccDBConnection("openDB")
    Set QS = ldb.QueryDefs("RSP_Ausw_Basisdaten_1_Export")
    If Not bFilterOn Then
        If InStr(QS.sql, "Where") Then
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, "Where") - 1) & Forms(sÜbersicht).Filter = sFilter & ";"   '    20180913-ROS
            QS.sql = Left(QS.sql, InStr(QS.sql, "Where") - 1) & sfilter & ";"
        End If
    Else
        If InStr(QS.sql, "Where") Then
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, "Where") + 5) & Forms(sÜbersicht).Filter = sFilter & ";"'    20180913-ROS
            QS.sql = Left(QS.sql, InStr(QS.sql, "Where") + 5) & sfilter & ";"
        Else
'            QS.SQL = Left(QS.SQL, InStr(QS.SQL, ";") - 1) & " Where " & Forms(sÜbersicht).Filter = sFilter & ";"   '    20180913-ROS
            QS.sql = Left(QS.sql, InStr(QS.sql, ";") - 1) & " Where " & sfilter & ";"
        End If
    End If

exit_Übernehmen_Click:
    Set rs = Nothing

    Forms(sÜbersicht).SetFocus
    Exit Sub

err_Übernehmen_Click:
    If Err = 3021 Then
        Resume Next
    Else
        GoTo exit_Übernehmen_Click
    End If
End Sub
Private Sub Schließen_Click()
On Error GoTo Err_Schließen_Click

    FormKoorSave Me
    DoCmd.Close

Exit_Schließen_Click:
    Exit Sub

Err_Schließen_Click:
    MsgBox Err.Description
    Resume Exit_Schließen_Click

End Sub
Private Sub Befehl107_Click()
On Error GoTo Err_Befehl107_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Befehl107_Click:
    Exit Sub

Err_Befehl107_Click:
    MsgBox Err.Description
    Resume Exit_Befehl107_Click

End Sub

Private Sub Von_Click()
    Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)
End Sub

Private Sub VorgangstypAlle_Click()
Me.ListVorgangstyp = Null
End Sub
