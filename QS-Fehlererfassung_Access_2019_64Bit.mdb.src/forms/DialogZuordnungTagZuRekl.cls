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
    Me.ListTagRekl.Requery
err_ButDeleteTag_Click:

    If Err = 3075 Then Exit Sub

End Sub

Private Sub ButDeleteTag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub ButNewTag_Click()

    Dim sTag As String
    Dim ldb As DAO.Database

    sTag = getValueFromInputBox(getTC_FCFromCode("IC3849"), getTC_FCFromCode("IC3850"))

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Insert into RSP_ReklTag (Tag) values ('" & sTag & "');")

    Me.ListTag.Requery

End Sub

Private Sub ButNewTag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)

    FormOpenKoorMousPointer Me, XPos, YPos, 630, 580, False

End Sub



Private Sub Schließen_Click()

    Dim l_Form As Form

    Set l_Form = m_parentDialog
    l_Form.closeInstanceForm l_Form

    Set l_Form = Nothing
    Forms("Fehlererfassung - Register").Controls("ListReklTag").Requery
    Forms("Fehlererfassung - Register").Controls("Liste6454").Requery


End Sub

Private Sub Schließen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub TagLoeschen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub TagZuordnen_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from ReklIDtoTagID where [ReklID] = " & m_parentDialog.Controls("Mastertab#ID") & " and [TagID] = " & Me.ListTag & ";")
    If rs.RecordCount = 0 Then
        ldb.Execute ("Insert into ReklIDtoTagID ([ReklID], [TagID]) values (" & m_parentDialog.Controls("Mastertab#ID") & ", " & Me.ListTag & ");")
    End If
    Set rs = Nothing
    Me.ListTagRekl.Requery

End Sub

Private Sub TagLoeschen_Click()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Delete * from ReklIDtoTagID where [ReklID] = " & m_parentDialog.Controls("Mastertab#ID") & " and [TagID] = " & Me.ListTagRekl & ";")
    Me.Requery
    Me.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

'########################################################
'################ Proberty Setting ######################
'########################################################

Public Property Let parentDialog(ByVal fValue As Form)

    Set m_parentDialog = fValue
    Me.ListTagRekl.RowSource = "SELECT ReklIDtoTagID.TagID, RSP_ReklTag.Tag FROM ReklIDtoTagID LEFT JOIN RSP_ReklTag ON ReklIDtoTagID.TagID = RSP_ReklTag.TagID " & _
                               "WHERE ReklIDtoTagID.ReklID = " & m_parentDialog.Controls("Mastertab#ID") & _
                               " GROUP BY ReklIDtoTagID.TagID, RSP_ReklTag.Tag;"
    Me.ListTagRekl.Requery


End Property

Public Property Get parentDialog() As Form

    Set parentDialog = m_parentDialog

End Property

Private Sub TagZuordnen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub
