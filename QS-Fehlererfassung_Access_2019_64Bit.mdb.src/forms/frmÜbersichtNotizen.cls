Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Sub setValue(ByVal CntDate As Date)

    Me.ErledigenBis = CntDate

End Sub
Private Sub chkAlleOffenenNotizen_AfterUpdate()

    setFilterÜbersichtNotizen

End Sub

Private Sub chkAlleOffenenNotizen_Click()

    setFilterÜbersichtNotizen

End Sub

Private Sub cmbAusgabeEcxel_Click()

    Dim strFilter As String
    Dim sql As String

    ' Filter nur übergeben, wenn er im Formular auch gerade aktiv angewendet wird
    If Me.FilterOn Then
        strFilter = Me.Filter
    Else
        strFilter = ""
    End If

     sql = Me.RecordSource
     sql = Replace(sql, ", tblNotizen.NotizFeldFrei;", "")

    ' Aufruf der globalen Prozedur mit Übergabe der Werte
    ' Optional können Sie als dritten Parameter einen eigenen Dateinamen mitgeben
    Call ExportRecordSourceNachExcel(sql, strFilter, Me.Name)

End Sub

Private Sub cmbAusgabeEcxel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbFilterLöschen_Click()
On Error GoTo Err_cmbFilterLöschen_Click

    Me.cboBearbKz = ""
    Me.chkAlleOffenenNotizen = 0
    Me.FilterOn = False
    Me.Refresh

Exit_cmbFilterLöschen_Click:
    Exit Sub

Err_cmbFilterLöschen_Click:
    MsgBox Err.Description
    Resume Exit_cmbFilterLöschen_Click

End Sub

Private Sub cboBearbKz_AfterUpdate()
    setFilterÜbersichtNotizen
End Sub



Private Sub cmbFilterLöschen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbOdrnerOeffnen_Click()

    Dim ReklnrKnd As String

    ReklnrKnd = Nz(DLookup("[Meldungsnummer Kunde]", "Mastertab", "[ID] = " & Me.ReklID & ""))
    Call VorgangsOrdner(Me.ReklID, ReklnrKnd, Me.ErledigenBis, True)   '20190820-ROS

End Sub

Private Sub cmbOdrnerOeffnen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub ErledigenBis_AfterUpdate()

    Me.Requery

End Sub

Private Sub ErledigenBis_Change()

    Me.Requery

End Sub

Private Sub ErledigenBis_Click()

    Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)

End Sub

Private Sub ErledigenBis_Dirty(Cancel As Integer)

    Me.Requery

End Sub

Private Sub ErledigenBis_LostFocus()

    Me.Requery

End Sub

Private Sub Erledigt_Click()

    If IsNull(Me.ErledigtAm) Then
        Me.ErledigtAm = Date
    Else
        Me.ErledigtAm = ""
    End If

    Me.Requery
    Me.Refresh

    txtAnzahlDatensatz = DCount("ReklID", "tblNotizen", "tblNotizen.ErledigenBis Is Not Null and Erledigt = False")

End Sub


Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)

    txtAnzahlDatensatz = DCount("ReklID", "tblNotizen", "tblNotizen.ErledigenBis Is Not Null and Erledigt = False")

End Sub

Private Sub Form_Close()
    FormKoorSave Me
End Sub

Private Sub Form_DblClick(Cancel As Integer)


    p_ReklNr = Me.ReklID

    Call NewReklamationE("Bearbeiten", get_p_ReklNr())

End Sub
Private Sub setFilterÜbersichtNotizen()

    Dim ls_Filter As String

    ls_Filter = ""

    If Me.chkAlleOffenenNotizen = -1 Then
        ls_Filter = ls_Filter & "(tblNotizen.Erledigt = False) AND "
    End If

    If Me.cboBearbKz <> "" Then
        ls_Filter = ls_Filter & "(BearbeiterKz = '" & Me.cboBearbKz & "')"
    End If

    If Right(ls_Filter, 5) = " AND " Then
        Me.Filter = Left(ls_Filter, Len(ls_Filter) - 5)
    Else
        Me.Filter = ls_Filter
    End If


    Me.FilterOn = True
    Me.Requery


End Sub

Private Sub Form_Load()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim WinAnmeldung As String
    Dim BearbKz As String

    WinAnmeldung = fOSUserName
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Name] = '" & WinAnmeldung & "';", dbOpenSnapshot)
        If rs.RecordCount Then
            BearbKz = rs("Kurzzeichen")
        Else
'            MsgBox ("Sacharbeiter ist nicht in Datenbank eingetragen")
            MsgBox getTC_FCFromCode("TC3213"), vbInformation, getTC_FCFromCode("TC3214")
            Exit Sub
        End If

    Me.chkAlleOffenenNotizen = 1
    Me.Filter = "Erledigt = False" 'and BearbeiterKz ='" & BearbKz & "'"
    Me.FilterOn = True
    txtAnzahlDatensatz = DCount("ReklID", "tblNotizen", "tblNotizen.ErledigenBis Is Not Null and Erledigt = False")

    Set rs = Nothing

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)
End Sub


Private Sub NotizFeld_AfterUpdate()

    Me.Refresh

End Sub
