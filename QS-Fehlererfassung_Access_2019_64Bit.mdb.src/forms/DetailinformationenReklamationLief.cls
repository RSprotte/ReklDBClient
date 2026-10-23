Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub BefundLiefIst_AfterUpdate()

    Me.[Rückmeldung_Lieferant_Verlängertl] = Me.Rückmeldung_Lieferant_Soll

End Sub

Private Sub ButOK_Click()
'20150115-RSP: Nullwerte bei beiden Parameter abgefangen
    DoCmd.RunCommand acCmdSaveRecord
    If IsNull(Me.BezBestellNrLQ) Or Me.BezBestellNrLQ = "" Then
        sBestellnummerLQ = ""
    Else
        sBestellnummerLQ = Me.BezBestellNrLQ
    End If
    If IsNull(Me.BezLieferscheinNrLieferant) Or Me.BezLieferscheinNrLieferant = "" Then
        sLieferscheinLieferant = ""
    Else
        sLieferscheinLieferant = Me.BezLieferscheinNrLieferant
    End If
    If IsNull(Me.BefundLiefIst) Or Me.BefundLiefIst = "" Then
        sReparaturdatum = ""
    Else
        sReparaturdatum = Me.BefundLiefIst
    End If


    DoCmd.Close acForm, "DetailinformationenReklamationLief"


End Sub

Private Sub cmbAbrechen_Click()

    DoCmd.Close

End Sub

Private Sub cmbReklamationsbericht_Click()

    Call ButOK_Click
    Forms("Fehlererfassung - Register").XDReportLief_DblClick (1)

End Sub

Private Sub cmbReklamationsbericht_E_Click()

    Call ButOK_Click
    Forms("Fehlererfassung - Register").XDReportLief_DblClick (2)

End Sub

Private Sub cmbRepAuftrag_Click()

    Call ButOK_Click
    Forms("Fehlererfassung - Register").ReparaturauftragExternErstellen

End Sub

Private Sub Form_Load()

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 420, 200

    Me.AnnahmeMitNacharbeit = 0
    Me.AnnahmeMitPrüfaufnahme = 0
    Me.RücklieferungWareLief = 0

End Sub
