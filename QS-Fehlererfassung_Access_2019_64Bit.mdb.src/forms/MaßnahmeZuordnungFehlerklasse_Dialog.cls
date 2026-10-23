Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Befehl9_Click()
End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

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

Private Sub Rechts_Click()
On Error GoTo Err_Rechts_Click

Dim ldb As DAO.Database
Set ldb = openAccDBConnection("openDB")

db.Execute ("Insert into MaßnahmenFehlerarten ([MaßnahmeID],[Fehlerklasse]) values (" & Me.MaßnahmeID & ", '" & Me.ListTaetigkeitenOhneMaßnahme & "');")
Me.Requery
Me.Refresh


 ' Set ldb = openAccDBConnection("closeDB")

Exit_Rechts_Click:
    Exit Sub

Err_Rechts_Click:
    MsgBox Err.Description
    Resume Exit_Rechts_Click


End Sub


Private Sub Links_Click()
On Error GoTo Err_Links_Click


Dim ldb As DAO.Database
Set ldb = openAccDBConnection("openDB")

db.Execute ("Delete * from MaßnahmenFehlerarten where (([MaßnahmeID]=" & Me.MID & ") and ([Fehlerklasse] ='" & Me.ListTaetigkeitenMaßnahme & "'));")
Me.Requery
Me.Refresh



Exit_Links_Click:
    Exit Sub

Err_Links_Click:
    MsgBox Err.Description
    Resume Exit_Links_Click


End Sub

Private Sub Schließen_Click()
On Error GoTo Err_Schließen_Click

    FormKoorSave Me
    DoCmd.Close
    Forms("MaßnahmenDialog").Controls("ListFehlerklassen").Requery

Exit_Schließen_Click:
    Exit Sub

Err_Schließen_Click:
    MsgBox Err.Description
    Resume Exit_Schließen_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
