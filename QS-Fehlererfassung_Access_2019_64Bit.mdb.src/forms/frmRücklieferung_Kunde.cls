Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit


Private Sub cmbDrucken_Click()

        Forms("frmRücklieferung_Kunde").Printer.Orientation = acPRORLandscape
        DoCmd.PrintOut acPrintAll

End Sub



Private Sub Form_DblClick(Cancel As Integer)

    If p_ReklNr <> "" Then
        unlockDataset (p_ReklNr)
    End If
    p_ReklNr = Me.ID
    NewReklamationE "Bearbeiten", p_ReklNr

End Sub

Private Sub cmbFrmSchließen_Click()
On Error GoTo Err_cmbFrmSchließen_Click

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * FROM   tblTempRückliefKunde;")
    ldb.Execute ("Update tblTempRückliefKunde set [Notiz] = '" & Me.txtNotizen & "' where ID = 1;")

    Set rs = Nothing

    FormKoorSave Me
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmbFrmSchließen_Click:
    Exit Sub

Err_cmbFrmSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbFrmSchließen_Click

End Sub

Private Sub Form_Load()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * FROM tblTempRückliefKunde where ID = 1;")

    Me.txtNotizen = rs("Notiz")

    Set rs = Nothing

    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub
