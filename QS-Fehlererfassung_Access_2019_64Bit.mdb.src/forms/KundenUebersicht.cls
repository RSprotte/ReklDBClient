Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AnsprechpartnerBearbeiten_Click()

Dim rsFormClone As DAO.Recordset

    plDialogSelectedAPID = Me.KundeAnsprechpartnerÜbersicht.Form.Controls("Kunden-ID")
    plDialogSelectedKundenID = Me.Firma_ID
    psDialogAnsprechpartnerKunde = "Bearbeiten"
    DoCmd.OpenForm "DialogAnsprechpartnerKunde", acNormal, , , , acDialog
    Forms("KundenUebersicht").SetFocus
    Me.Requery
    Set rsFormClone = Me.RecordsetClone
    With rsFormClone
        .FindFirst "Firma_ID=" & plDialogSelectedKundenID
        If Not .NoMatch Then
            Me.Bookmark = .Bookmark
        End If
        .Close
    End With

    Set rsFormClone = Nothing

End Sub

Private Sub AnsprechpartnerLoeschen_Click()

    Dim ldb As DAO.Database
    Dim rsFormClone As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    plDialogSelectedAPID = Me.KundeAnsprechpartnerÜbersicht.Form.Controls("Kunden-ID")
    ldb.Execute ("update Kunden set [Aktiv] = False where [Kunden-ID] = " & Me.KundeAnsprechpartnerÜbersicht.Form.Controls("Kunden-ID") & ";")
    Me.KundeAnsprechpartnerÜbersicht.Form.Requery

    Set rsFormClone = Me.KundeAnsprechpartnerÜbersicht.Form.RecordsetClone
    With rsFormClone
        .FindFirst "[Kunden-ID]=" & plDialogSelectedAPID
        If Not .NoMatch Then
            Me.KundeAnsprechpartnerÜbersicht.Form.Bookmark = .Bookmark
        End If
        .Close
    End With

    Set rsFormClone = Nothing

End Sub

Private Sub Befehl7_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub



Private Sub FilterEingabe_Change()      '20190823-ROS
On Error Resume Next

    Dim strFilter As String
    Dim intStart As Integer
    Dim i As String

    intStart = Me.FilterEingabe.SelStart

    If Me.OptionFilter = 3 Then
        Me.KundeAnsprechpartnerÜbersicht.LinkMasterFields = ""
        Me.KundeAnsprechpartnerÜbersicht.LinkChildFields = ""
        strFilter = "[Kontaktperson] like '" & Me.FilterEingabe.Text & "*'"
        Me.KundeAnsprechpartnerÜbersicht.Form.FilterOn = False
        Me.KundeAnsprechpartnerÜbersicht.Form.Filter = strFilter
        Me.KundeAnsprechpartnerÜbersicht.Form.FilterOn = True
        Me.FilterEingabe.SelStart = intStart
        Me.FilterEingabe.SetFocus
    Else
        If Not Len(Me.FilterEingabe.Text) = 0 Then
            If Me.OptionFilter = 1 Then
                strFilter = "[KndNr_Avista] like '" & Me.FilterEingabe.Text & "*'"
            End If
            If Me.OptionFilter = 2 Then
                strFilter = "[Firma_Name] like '" & Me.FilterEingabe.Text & "*'"
            End If
            Me.Filter = strFilter
            Me.FilterOn = True
            Me.FilterEingabe.SelStart = intStart
            Me.FilterEingabe.SetFocus
        Else
            Me.FilterEingabe = ""
            Me.FilterOn = False
            Me.FilterEingabe.SetFocus
        End If
        Me.KundeAnsprechpartnerÜbersicht.Form.FilterOn = False
        Me.KundeAnsprechpartnerÜbersicht.LinkMasterFields = "Firma_ID"
        Me.KundeAnsprechpartnerÜbersicht.LinkChildFields = "FirmaID"
    End If

End Sub


Private Sub FilterLöschen_Click()
'    Dim sfilter As String
    Me.KundeAnsprechpartnerÜbersicht.LinkMasterFields = "Firma_ID"
    Me.KundeAnsprechpartnerÜbersicht.LinkChildFields = "FirmaID"
    With Me.KundeAnsprechpartnerÜbersicht.Form
        .OrderBy = "[Kontaktperson] ASC"
        .OrderByOn = True
    End With
    Me.FilterEingabe = ""
    Me.FilterOn = False
    Me.KundeAnsprechpartnerÜbersicht.Form.FilterOn = False

'    sfilter = ""
'    Me.Filter = ""
'    Me.FilterOn = True
'    Me.FilterEingabe = ""
'    Me.Filter = sfilter
'    Me.Requery
'    Me.Refresh

End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

    Me.OptionFilter = 2   '20190828_ROS_1
    Me.FilterEingabe.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub


Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
'    Call FormEinUndAusBlenden(Me)
    Me.FilterEingabe.SetFocus

'    Dim XPos As Long
'    Dim YPos As Long
'    Dim lgPosition As Long
'
'    lgPosition = getPointerPos(XPos, YPos)
'    FormOpenKoorMousPointer Me, XPos - 500, YPos - 100, 800, 680, False

End Sub



Private Sub NeuerAnsprechpartner_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim bNewAP As Boolean
    Dim pDsAP As Long
    Dim rsFormClone As DAO.Recordset
    Dim rsClone As DAO.Recordset

    plDialogSelectedKundenID = Me.Firma_ID
    psDialogAnsprechpartnerKunde = "Neu"
    DoCmd.OpenForm "DialogAnsprechpartnerKunde", acNormal, , , , acDialog
    Forms("KundenUebersicht").SetFocus
    Me.Requery
    Set rsFormClone = Me.RecordsetClone
    With rsFormClone
        .MoveLast
        .FindFirst "Firma_ID=" & plDialogSelectedKundenID
        If Not .NoMatch Then
            Me.Bookmark = .Bookmark
        End If
        .Close
    End With

    With Me.KundeAnsprechpartnerÜbersicht.Form
        .OrderBy = "[Kunden-ID] DESC"
        .OrderByOn = True
    End With

    Set rsFormClone = Nothing

End Sub

Private Sub NeuerKunde_Click()

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim lAvistaKD As String
    Dim sql As String

'    psDialogKunde = "NeuerKunde"
'    DoCmd.OpenForm "DialogKunde", acNormal, , , , acDialog

' 20260724-ROS: Kunden anlegen vereinfacht wegen Einführung der neuen Kundengruppen im AVISTA

    'Eingabe der Avista-Kundennummer
    lAvistaKD = InputBox(getTC_FCFromCode("IC4293"), getTC_FCFromCode("IC4294"))

    'sql-String zum suchen der Kundennummer in AVISTA
    sql = "SELECT DISTINCT ta_kundengr_bez.kdgrp_id, tm_partei.pa_org_name1 " & _
            "FROM DBA.tm_partei INNER JOIN DBA.tm_kunde ON tm_partei.pa_id =  tm_kunde.adk_id " & _
            "INNER JOIN DBA.ta_kundengr  ON ta_kundengr.kdgrp_id =  tm_kunde.kdgrp_id " & _
            "LEFT OUTER JOIN DBA.ta_kundengr_bez ON tm_kunde.kdgrp_id = ta_kundengr_bez.kdgrp_id " & _
            "WHERE tm_kunde.adk_nummer = '##KD##' and ta_kundengr_bez.sprach_id = 1"

    Set ldb = openAccDBConnection("openDB")

    'Suchen der Kundennummer in Datenbank
    Set rs = ldb.OpenRecordset("Select * from tblFirma where [KndNr_Avista] = '" & lAvistaKD & "';", dbOpenSnapshot)

    'Wenn vorhanden wird Abgebrochen, wenn nicht wird neu angelegt mit den Daten aus AVISTA
    If rs.RecordCount Then
        MsgBox getTC_FCFromCode("TC4295"), vbInformation, getTC_FCFromCode("TC4296")
        Set rs = Nothing
        Exit Sub
    Else
        'Suchen der benötigten Daten in AVISTA
        sql = Replace(sql, "##KD##", lAvistaKD)
        Set adbprov = New clsAvistaProvider
        Set rsAvista = New ADODB.Recordset
        Set rsAvista = adbprov.getRecordsetAvista(sql)
        If rsAvista.RecordCount Then
            ldb.Execute ("Insert into tblFirma ([Firma_Name],[Zurodnung_ID], [KndNr_Avista]) values ('" & rsAvista("pa_org_name1") & "', '" & rsAvista("kdgrp_id") & "','" & lAvistaKD & "');")
        Else
            MsgBox getTC_FCFromCode("TC4297"), vbInformation, getTC_FCFromCode("TC4296")
            Set rsAvista = Nothing
            Exit Sub
        End If
        Set rsAvista = Nothing
    End If

    'Übersicht aktualisieren und Fokus auf neu angelegten Kunden setzten
    Forms("KundenUebersicht").SetFocus
    Me.Requery
    Me.Refresh
    DoCmd.GoToRecord acDataForm, Me.Name, acLast

    Set rs = Nothing

End Sub

Private Sub KundeBearbeiten_Click()

    Dim liKndID As Integer
    Dim rsFormClone As DAO.Recordset

    psDialogKunde = "BearbeitenKunde"
    liKndID = Me.Firma_ID
    DoCmd.OpenForm "DialogKunde", acNormal, , , , acDialog

    Me.Requery
    Me.Refresh
    Set rsFormClone = Me.RecordsetClone
    With rsFormClone
        .FindFirst "Firma_ID=" & liKndID
        If Not .NoMatch Then
            Me.Bookmark = .Bookmark
        End If
        .Close
    End With

    Set rsFormClone = Nothing

End Sub

Private Sub Objekt_schließen___Formular_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub
