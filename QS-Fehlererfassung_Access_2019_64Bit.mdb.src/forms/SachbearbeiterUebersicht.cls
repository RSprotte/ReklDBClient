Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Befehl7_Click()

    FormKoorSave Me
    Call FilterLöschen_Click
    DoCmd.Close

End Sub

Private Sub Befehl8_Click()
Dim ldb As DAO.Database

 Set ldb = openAccDBConnection("openDB")

If MsgBox(getTC_FCFromCode("TC0227", ldb), vbOKCancel) = vbOK Then  'TC0227
    ldb.Execute ("delete * from Sachbearbeiter where Name = '" & Me.NameSachbearbeiter & "';")
    Me.Requery
    Me.Refresh
End If

 ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub Befehl9_Click()

DoCmd.OpenForm "Sachbearbeiter"
With Forms("Sachbearbeiter")
    .Filter = "[SB-ID] = " & Me.SB_ID & " "
    .FilterOn = True
    .Requery
    .Refresh
End With

'AllowAdditions = False

End Sub

Private Sub FilterEinheit_AfterUpdate()

    Me.Filter = "[Einheit] = '" & Me.FilterEinheit & "'"
    Me.FilterOn = True
    Me.Requery
    Me.Refresh

End Sub

Private Sub FilterLöschen_Click()


    Me.Filter = ""
    Me.FilterOn = True
    Me.FilterEinheit = ""
    Me.Requery
    Me.Refresh



End Sub

Private Sub Form_DblClick(Cancel As Integer)

Befehl9_Click

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

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Neu_Click()

DoCmd.OpenForm "Sachbearbeiter"
DoCmd.GoToRecord , , acNewRec

AllowAdditions = False

End Sub
