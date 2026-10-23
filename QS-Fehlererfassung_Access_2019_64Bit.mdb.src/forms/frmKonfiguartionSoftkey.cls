Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdDel_Click(fIndex As Integer)


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeyS where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";")
    If rs.RecordCount Then
        rs.Delete
    End If
    Me.Controls("txtMakroSoftKey" & fIndex) = ""
    Me.Controls("txtMakroSoftKeyBez" & fIndex) = ""

    If IsLoaded("TB:autostart") Then
        If Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = True Then
            Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = False
        End If
    End If

    Set rs = Nothing


End Sub

Private Sub cmdDel1_Click()
    cmdDel_Click 1
End Sub
Private Sub cmdDel2_Click()
    cmdDel_Click 2
End Sub
Private Sub cmdDel3_Click()
    cmdDel_Click 3
End Sub
Private Sub cmdDel4_Click()
    cmdDel_Click 4
End Sub
Private Sub cmdDel5_Click()
    cmdDel_Click 5
End Sub


Private Sub cmdSet_Click(fIndex As Integer)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeyS where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";")
    If rs.RecordCount Then
        rs.Edit
    Else
        rs.AddNew
        rs("SoftkeyNummer") = fIndex
        rs("SchbearbeiterKuerzel") = get_pKuerzel()

    End If
    rs("SoftkeyMakro") = Me.lstSoftKeyMakroZuordnung.Column(1)
    Me.Controls("txtMakroSoftKey" & fIndex) = Me.lstSoftKeyMakroZuordnung.Column(1)
    rs("SoftKexBezeichnung") = Me.lstSoftKeyMakroZuordnung.Column(2)
    Me.Controls("txtMakroSoftKeyBez" & fIndex) = Me.lstSoftKeyMakroZuordnung.Column(2)
    rs.Update
    'me.

    If IsLoaded("TB:autostart") Then
        If IsNull(Me.Controls("txtMakroSoftKeyBez" & fIndex)) Then
'            MsgBox ("Kein Soft Key ausgewählt")
            MsgBox getTC_FCFromCode("TC3199"), vbInformation, getTC_FCFromCode("TC3200")
            Set rs = Nothing
            Set db = Nothing
            Exit Sub
        End If
        Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Caption = Me.Controls("txtMakroSoftKeyBez" & fIndex)
        If Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = False Then
            Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = True
        End If
    End If

    Set rs = Nothing
    Set db = Nothing


End Sub
Private Sub cmdSet1_Click()

    cmdSet_Click 1

End Sub
Private Sub cmdSet2_Click()

    cmdSet_Click 2

End Sub
Private Sub cmdSet3_Click()

    cmdSet_Click 3

End Sub
Private Sub cmdSet4_Click()

    cmdSet_Click 4

End Sub
Private Sub cmdSet5_Click()

    cmdSet_Click 5

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

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

Private Sub Form_Open(Change As Integer)
On Error Resume Next


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblSoftkeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' order by SoftkeyNummer;")

        While Not rs.EOF
            Me.Controls("txtMakroSoftKey" & rs("SoftkeyNummer")) = rs("SoftkeyMakro")
            Me.Controls("txtMakroSoftKeyBez" & rs("SoftkeyNummer")) = rs("SoftKexBezeichnung")
            rs.MoveNext
        Wend

    Set rs = Nothing
    Set db = Nothing


End Sub

Private Sub Schließen_Click()
FormKoorSave Me
DoCmd.Close

End Sub

Private Sub txtMakroSoftKeyBez_AfterUpdate(fIndex As Integer)
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")

    If Me.Controls("txtMakroSoftKeyBez" & fIndex).Text = "" Then
'        MsgBox "Leere Zeichenfolge ist nicht erlaubt!"
        MsgBox getTC_FCFromCode("TC4234"), vbInformation, getTC_FCFromCode("TC4235")
    Else
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from tblSoftKeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";")
        If rs.RecordCount Then
            rs.Edit
            rs("SoftKexBezeichnung") = Me.Controls("txtMakroSoftKeyBez" & fIndex).Text
            rs.Update
            If IsLoaded("TB:autostart") Then
                Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Caption = Me.Controls("txtMakroSoftKeyBez" & fIndex).Text
            End If
        End If
        Set rs = Nothing
    End If
End Sub

Private Sub txtMakroSoftKeyBez1_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 1

End Sub
Private Sub txtMakroSoftKeyBez2_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 2

End Sub
Private Sub txtMakroSoftKeyBez3_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 3

End Sub
Private Sub txtMakroSoftKeyBez4_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 4

End Sub
Private Sub txtMakroSoftKeyBez5_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 5

End Sub
