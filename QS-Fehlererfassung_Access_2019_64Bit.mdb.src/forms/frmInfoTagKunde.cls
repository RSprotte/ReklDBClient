Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim m_parentDialog As Form

Private Sub ButDeleteTag_Click()
On Error GoTo err_ButDeleteTag_Click

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    If MsgBox(getTC_FCFromCode("TC4216") & " - " & Me.ListTag.Column(1) & " - " & getTC_FCFromCode("TC4217"), vbYesNo + vbDefaultButton1) = vbYes Then
        ldb.Execute ("delete * from ReklIDtoTagID where TagID = " & Me.ListTag & ";")
        ldb.Execute ("delete * from RSP_ReklTag where TagID = " & Me.ListTag & ";")
    End If
    Me.ListTag.Requery

err_ButDeleteTag_Click:

    If Err = 3075 Then Exit Sub

End Sub

Private Sub ButDeleteTag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub ButNewTag_Click()

    Dim sTag As String
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    sTag = Me.cboFirma.Column(2) & " / " & getValueFromInputBox(getTC_FCFromCode("IC3849"), getTC_FCFromCode("IC3850"))

'    sTag = Me.cboFirma.Column(2) & " / " & InputBox(getTC_FCFromCode("IC3849"), getTC_FCFromCode("IC3850"))

    ldb.Execute ("Insert into RSP_ReklTag (Tag) values ('" & sTag & "');")
    Me.ListTag.Requery

End Sub

Private Sub ButNewTag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Open(Cancel As Integer)

    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub



Private Sub Schließen_Click()

    FormKoorSave Me
    DoCmd.Close acForm, Me.Name

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub



Private Sub Schließen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub
