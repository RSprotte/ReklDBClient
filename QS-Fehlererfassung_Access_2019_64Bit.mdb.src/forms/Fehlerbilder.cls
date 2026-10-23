Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub FehlerbilderAnalyse_Click()

    DoCmd.OpenForm "FehlerbilderAnalyse", acNormal
    With Forms("FehlerbilderAnalyse")
        .Controls("ListFehlerbilder") = Me.[Einzeltaetigkeit-ID]
        .Controls("ListFehlerbilder").Requery
        .Requery
        .Refresh
        .Controls("ListFehlerbilder").Locked = True
        .Controls("FehlerbildAnalyse").Locked = True
        .UF01.Controls("Beschreibung").Locked = True
        .Controls("Übernehmen").Enabled = False
        Forms("Fehlerbilder").AllowAdditions = False
    End With


End Sub

Private Sub Form_Load()

'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub


Private Sub Neu_Click()

        DoCmd.OpenForm "Dialog_DefTaetigkeit"
        DialogDefTaetigkeit = "Neu"
        TaetigkeitID = Me.ListTaetigkeiten

End Sub
Private Sub Schließen_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
