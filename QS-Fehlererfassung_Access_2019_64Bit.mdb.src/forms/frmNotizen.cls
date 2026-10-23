Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim sOpenArgs As String
Dim aArray As Variant
Public Sub setValue(ByVal CntDate As Date)

    Me.ErledigenBis = CntDate

End Sub

Private Sub ErledigenBis_Click()

    Dim frm As Form

    DoCmd.OpenForm ("frmCalendar")
    Set frm = Forms("frmCalendar")
    Call frm.setParentForm(Me)

End Sub

Private Sub Form_Load()

'Dim sOpenArgs As String
'Dim aArray As Variant
Dim i As Integer

    sOpenArgs = Me.OpenArgs
    aArray = Split(sOpenArgs, ",")

    For i = 0 To (UBound(aArray) - i)
    '    MsgBox aArray(i)
    Next
        Me.ReklID = aArray(0)
    '    Me.StatusVorgang = aArray(1)
        Me.BearbeiterKz = aArray(2)


End Sub
Private Sub cmbfrmNotizSchließen_Click()
On Error GoTo Err_cmbfrmNotizSchließen_Click


    If IsNull(Me.ErledigenBis) Then
'        MsgBox ("Bitte Datum eintragen")
        MsgBox getTC_FCFromCode("TC3211"), vbInformation, getTC_FCFromCode("TC3212")
        Exit Sub
    End If

    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close acForm, "frmNotizen", acSaveYes

    If CurrentProject.AllForms("frmÜbersichtNotizen").IsLoaded Then
        Forms("frmÜbersichtNotizen").Requery
        Forms("frmÜbersichtNotizen").Refresh
    End If

    If IsLoaded("frmFormularROS") Then
        With Forms("frmFormularROS")
            .Refresh
            .Form("frmÜbersichtNotizen").Controls("txtAnzahlDatensatz") = DCount("ReklID", "tblNotizen", "tblNotizen.ErledigenBis Is Not Null and Erledigt = False")
        End With
    End If

Exit_cmbfrmNotizSchließen_Click:
    Exit Sub

Err_cmbfrmNotizSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbfrmNotizSchließen_Click

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long


    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 580, 540


End Sub

Private Sub NotizFeld_AfterUpdate()

    sOpenArgs = Me.OpenArgs
    aArray = Split(sOpenArgs, ",")

    Me.NotizFeldFrei = Me.NotizFeld.Column(1)
    Me.NotizFeldFrei = aArray(3) & ": " & Me.NotizFeld.Column(1)


End Sub
Private Sub cmbFromSchließen_Click()
On Error GoTo Err_cmbFromSchließen_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
    DoCmd.Close

Exit_cmbFromSchließen_Click:
    Exit Sub

Err_cmbFromSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbFromSchließen_Click

End Sub
