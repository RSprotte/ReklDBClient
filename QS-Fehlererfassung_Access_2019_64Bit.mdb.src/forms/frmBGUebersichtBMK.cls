Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public lexcelprov As clsExcelProvider
Public adbprov As clsAvistaProvider

'Formularverwaltung
'Private mclsFormClass As clsFormClass



Private Sub cmbGetAvista_Click()

    Dim ls_currArtikelBMK As String
    Dim rsAvista As ADODB.Recordset
    Dim rsAvistaPos As ADODB.Recordset
    Dim ls_sql As String
    Dim ls_sql_STLID As String
    Dim ls_sql_STL As String
    Dim ListArray(2, 1, 1) As String
    Dim ls_Values As String
    Dim ls_ValuesBG As String
    Dim fld As Field
    Dim lr_FktRange As Range
    Dim lr_BMKRange As Range
    Dim lr_BMK As Range
    Dim ls_BMKAddress As String

    DoCmd.Hourglass True
    ls_currArtikelBMK = Me.txtBMK

    Me.txtBMK = ls_currArtikelBMK
    '###################################################################################
    'Ermitteln der Baugruppen in Avista anhand des angezeigten BMK
    '###################################################################################
    If Me.txtBMK <> "" Then
        'Avista DB Provider Instanz erzeugen
        If adbprov Is Nothing Then Set adbprov = New clsAvistaProvider

        If adbprov.genTempTableBGUbersichtBMK(Me.txtBMK, "tblTempBMK_BG_Pos_Uebersicht", Me.txtMachNr) <> "" Then
            MsgBox "Fehler in genTempTableBGUbersichtBMK()"
        End If


    End If
    Me.lstBGUebersicht.Requery
    DoCmd.Hourglass False
    Set adbprov = Nothing

End Sub

Private Sub cmbStueliOeffnen_Click()

    Dim ls_sql As String

    If IsNull(lstBGUebersicht.Column(1)) Then
        Exit Sub
    Else
        ls_sql = cStueli
        ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", AVISTA_ID_KDnr(Forms("Fehlererfassung - Register").Controls("Mastertab#ID")))
        ls_sql = Replace(ls_sql, "###Art###", lstBGUebersicht.Column(1))
        Call Stueli_Auslesen(ls_sql, lstBGUebersicht.Column(1))
    End If

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub



Private Sub Form_Load()

    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me
    CurrentDb().Execute ("delete * from tblTempBMK_BG_Pos_Uebersicht")
    Me.lstBGUebersicht.Requery
    Set RefCurrNode = Nothing
    Me.txtMachNr = ""
    Me.txtBMK = ""

End Sub

Private Sub lstBGUebersicht_Click()

    If IsLoaded("frmStueli") Then
        Call cmbStueliOeffnen_Click
    End If

End Sub

Private Sub txtMachNr_AfterUpdate()

    If IsNull(Me.txtMachNr) Then Me.txtMachNr = ""

End Sub

Private Sub cmbSchließen_Click()
On Error GoTo Err_cmbSchließen_Click

    If IsLoaded("frmStueli") Then
     DoCmd.Close acForm, "frmStueli"
    End If
    Me.SetFocus

    FormKoorSave Me
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmbSchließen_Click:
    Exit Sub

Err_cmbSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbSchließen_Click

End Sub
