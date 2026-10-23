Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub lstControlAktivForm_AfterUpdate()

    Me.frmUFControlsSteuerung.Form.FilterOn = False
    Me.frmUFControlsSteuerung.Form.Filter = "[ControlName] = '" & Me.lstControlAktivForm.Value & "'"
    Me.frmUFControlsSteuerung.Form.FilterOn = True

End Sub

Private Sub Objekt_schließen___Formular_Click()

    DoCmd.Close

End Sub


Public Sub Init_Form_Cmb(fForm As Form, fActiveControl As Object, fActivePage As Page)

    Dim i As Boolean
    Dim lIndxI As Integer
    Dim lIndxA As Integer
    Dim sTemp As String

    'Füllen aller Controls aus der ControlList (object.controls)
    If Not fActivePage Is Nothing Then i = AddControlListPage(fForm, fActivePage, Me.lstControlAktivForm)
    i = AddControlListForm(fForm, Me.lstControlAktivForm)
    'Aktives Control anzeigen
    Me.txtAktivControll = fActiveControl.Name
    Me.txtAktivFormName = fForm.Name
    txtAktivControll_AfterUpdate
    Me.txtAktivControll.SetFocus
    Me.Caption = "Control-Steuerung: " & fForm.Name

    ListeSortieren Me.lstControlAktivForm

End Sub
Private Sub cmbAktivControllSpeichern_Click()
On Error GoTo Err_cmbAktivControllSpeichern_Click


    DoCmd.RunCommand acCmdSaveRecord

Exit_cmbAktivControllSpeichern_Click:
    Exit Sub

Err_cmbAktivControllSpeichern_Click:
    MsgBox Err.Description
    Resume Exit_cmbAktivControllSpeichern_Click

End Sub

Private Sub txtAktivControll_AfterUpdate()

    Me.frmUFControlsSteuerungAktiv.Form.FilterOn = False
    Me.frmUFControlsSteuerungAktiv.Form.Filter = "[ControlName] = '" & Me.txtAktivControll.Value & "'"
    Me.frmUFControlsSteuerungAktiv.Form.FilterOn = True

End Sub
