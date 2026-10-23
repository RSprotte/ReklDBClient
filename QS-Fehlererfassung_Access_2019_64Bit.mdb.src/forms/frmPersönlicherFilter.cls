Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbFilterBearbeiten_Click()

    Dim AntwortEigeneMsgBox As Variant
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblFilterForm where [FilterFormID] = " & Me.FilterFormID & ";")

'    AntwortEigeneMsgBox = EigeneMsgBox("Bearbeitung auswählen", vbYesNoCancel, "Filter bearbeiten")
    AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3836"), vbYesNoCancel, getTC_FCFromCode("TC3837"))

    If AntwortEigeneMsgBox = vbYes Then
        If rs("Aktiv") = True Then
            rs.Edit
            rs("Aktiv") = False
            rs.Update
        Else
            rs.Edit
            rs("Aktiv") = True
            rs.Update
        End If
    End If
    If AntwortEigeneMsgBox = vbNo Then
        ldb.Execute ("Delete * from tblFilterForm where [FilterFormID] = " & Me.FilterFormID & ";")
    End If
    If AntwortEigeneMsgBox = vbCancel Then
        Exit Sub
    End If

    Me.Refresh
    DoCmd.RunCommand acCmdSelectRecord

    Set rs = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown KeyCode, Shift, Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyUp KeyCode, Shift
End Sub

Private Sub Form_Load()

'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

    Dim objMouseOverForm As clsMouseoverForm

    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub cmbClose_Click()

    FormKoorSave Me
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

End Sub
