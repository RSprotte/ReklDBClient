Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()
    If IsLoaded("SachbearbeiterUebersicht") Then
        Forms("SachbearbeiterUebersicht").Requery
        Forms("SachbearbeiterUebersicht").Refresh
        'me.[Sachbearbeiter].Form![Objekt schließen - Formular].Visible = False
    End If

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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Open(Cancel As Integer)
'MsgBox ("Öffnen")
If IsLoaded("SachbearbeiterUebersicht") Then
    Me.Schließen.Visible = True
    Me.NavigationButtons = False
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
